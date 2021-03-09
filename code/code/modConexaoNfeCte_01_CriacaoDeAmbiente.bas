Attribute VB_Name = "modConexaoNfeCte_01_CriacaoDeAmbiente"
Option Compare Database

Private Comando As Variant
Private con As ADODB.Connection

Private Const qryOrigemDestino As String = "INSERT INTO tblOrigemDestino (Destino,Tipo) VALUES('strDestino','strTipo')"
Private Const qryParametro As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"

Private Const deleteProcessamento As String = "DROP TABLE tblProcessamento"
Private Const Processamento As String = "CREATE TABLE tblProcessamento(ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,pk TEXT (50),chave TEXT (255),valor TEXT (255))"

Private Const deleteDados As String = "DROP TABLE tblDadosConexaoNFeCTe"
Private Const createDados As String = "CREATE TABLE tblDadosConexaoNFeCTe(ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,ID_Empresa Integer,ID_Tipo Integer,codMod Integer,codIntegrado Integer,dhEmi TEXT (50),CNPJ_emit TEXT (50),Razao_emit TEXT (255),CNPJ_Rem TEXT (50),CPNJ_Dest TEXT (50),CaminhoDoArquivo TEXT (255),Chave TEXT (255),Comando TEXT (255));"

Private Const deleteOrigemDestino As String = "DROP TABLE tblOrigemDestino"
Private Const createOrigemDestino As String = "CREATE TABLE tblOrigemDestino (ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY , Origem TEXT(255), Destino TEXT(255), Tipo TEXT(255))"

Private Const deleteParametros As String = "DROP TABLE tblParametros"
Private Const createParametros As String = "CREATE TABLE tblParametros (ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,TipoDeParametro TEXT(50), ValorDoParametro TEXT (255));"

Private Const deleteTipos As String = "DROP TABLE tblTipos"
Private Const createTipos As String = "CREATE TABLE tblTipos (ID AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,codMod Integer,Descricao TEXT (255));"

Private Const deleteCompras As String = "DROP TABLE tblCompraNF"
Private Const createCompras As String = "CREATE TABLE tblCompraNF (ID_CompraNF  AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY,IDOLD_CompraNF Integer,Fil_CompraNF  TEXT (255),NumNF_CompraNF  TEXT (255),NumPed_CompraNF  TEXT (255),NumOrc_CompraNF  TEXT (255),Esp_CompraNF  TEXT (255),Serie_CompraNF  TEXT (255),TPNF_CompraNF  TEXT (255),ID_NatOp_CompraNF  Integer,ID_NatOpOLD_CompraNF  Integer,CFOP_CompraNF  TEXT (255),IESubsTrib_CompraNF  TEXT (255),DTEmi_CompraNF  date,DTEntd_CompraNF  date,HoraEntd_CompraNF  TEXT (255),ID_Forn_CompraNF  Integer,ID_FornOld_CompraNF  Integer,ID_Compr_CompraNF  Integer,ID_Transp_CompraNF  Integer,ID_CondPgto_CompraNF  Integer,BaseCalcICMSSubsTrib_CompraNF  double,VTotICMSSubsTrib_CompraNF  double,VTotFrete_CompraNF  double,VTotSeguro_CompraNF  double,VTotOutDesp_CompraNF  double,BaseCalcICMS_CompraNF  double,VTotICMS_CompraNF  double,VTotIPI_CompraNF  double,VTotISS_CompraNF  double,VTotProd_CompraNF  double,VTotServ_CompraNF  double," & _
                                        "VTotNF_CompraNF  double,TxDesc_CompraNF  double,VTotDesc_CompraNF  double,TPFrete_CompraNF  double,Placa_CompraNF  TEXT (255),UFVeic_CompraNF  TEXT (255),QtdVol_CompraNF  Integer,EspVol_CompraNF TEXT (255),MarcaVol_CompraNF  TEXT (255),NumVol_CompraNF  TEXT (255),PesoBrt_CompraNF  double,PesoLiq_CompraNF  double,DdAdic_CompraNF  TEXT (255),Obs_CompraNF  TEXT (255),Sit_CompraNF  TEXT (255),IDCli_Depto_CompraNF  Integer,IDCli_Contato_CompraNF  Integer,IDCli_Email_CompraNF  Integer,IDCli_Fone_CompraNF  Integer,CondEsp_CompraNF  TEXT (255),Validade_CompraNF  TEXT (255),PzEntg_CompraNF TEXT (255)," & _
                                        "Garantia_CompraNF  TEXT (255),FlagSimples_CompraNF  TEXT (255),FlagDescBaseICMS_CompraNF  TEXT (255),FlagExp_CompraNF  TEXT (255),ModeloDoc_CompraNF  TEXT (255),ChvAcesso_CompraNF  TEXT (255),VTotPIS_CompraNF  double,VTotCOFINS_CompraNF  double,VTotPISRet_CompraNF  double,VTotCOFINSRet_CompraNF  double,VTotCSLLRet_CompraNF  double,VTotIRRFRet_CompraNF  double,FlagSomaST_CompraNF  TEXT(255),FlagCalculado_CompraNF  TEXT (255),VTotISSRet_CompraNF  double,DTExt_CompraNF  date,CNPJ_CPF_CompraNF  TEXT (255),NomeCompleto_CompraNF  TEXT (255),NomeCompleto_VendaNF  TEXT (255),ID_Imp_CompraNF  Integer,SitOR_CompraNF  TEXT (255),NumOR_CompraNF  TEXT (255),FlagNEnvWMAS_CompraNF  TEXT (255),IDVD_CompraNF  TEXT (255),IDVendaNF_CompraNF  TEXT (255),FlagTransf_CompraNF  TEXT(255));"

Private Const deleteComprasItens As String = "DROP TABLE tblCompraNFItem"
Private Const createComprasItens As String = "CREATE TABLE tblCompraNFItem (ID_CompraNFItem  AutoIncrement CONSTRAINT PrimaryKey PRIMARY KEY , " & _
                                            "IDOLD_CompraNFItem Integer , ID_CompraNF_CompraNFItem Integer , ID_CompraNFOLD_CompraNFItem Integer , Item_CompraNFItem Integer , ID_Prod_CompraNFItem Integer , ID_ProdOld_CompraNFItem Integer , ID_Grade_CompraNFItem Integer , Almox_CompraNFItem Integer , QtdFat_CompraNFItem Integer , VUnt_CompraNFItem double , TxDesc_CompraNFItem double , VUntDesc_CompraNFItem double , ICMS_CompraNFItem double , ISS_CompraNFItem double , IPI_CompraNFItem double , ID_NatOp_CompraNFItem Integer , ID_NatOpOLD_CompraNFItem Integer , CFOP_CompraNFItem TEXT (255) , ST_CompraNFItem TEXT (255) , FlagEst_CompraNFItem Byte , EstDe_CompraNFItem TEXT (255) , EstPara_CompraNFItem TEXT (255) , DTEmi_CompraNFItem date , Esp_CompraNFItem TEXT (255) , Série_CompraNFItem TEXT (255) , Num_CompraNFItem TEXT (255) , Dia_CompraNFItem TEXT (255) , UF_CompraNFItem TEXT (255) , VTot_CompraNFItem double , " & _
                                            "VCntb_CompraNFItem double , BaseCalcICMS_CompraNFItem double , VTotBaseCalcICMS_CompraNFItem double , DebICMS_CompraNFItem double , IseICMS_CompraNFItem double , OutICMS_CompraNFItem double , BaseCalcIPI_CompraNFItem double , DebIPI_CompraNFItem double , IseIPI_CompraNFItem double , OutIPI_CompraNFItem double , Obs_CompraNFItem TEXT (255) , TxMLSubsTrib_CompraNFItem double , TxIntSubsTrib_CompraNFItem double , TxExtSubsTrib_CompraNFItem double , BaseCalcICMSSubsTrib_CompraNFItem double , VTotICMSSubsTrib_compranfitem double , VTotDesc_CompraNFItem double , VTotFrete_CompraNFItem double ,  VTotSeg_CompraNFItem double , STIPI_CompraNFItem TEXT (255) , STPIS_CompraNFItem TEXT (255) , STCOFINS_CompraNFItem TEXT (255) , nID_CompraNFItem TEXT (255) , PIS_CompraNFItem double ,  COFINS_CompraNFItem double , VTotBaseCalcPIS_CompraNFItem double , VTotBaseCalcCOFINS_CompraNFItem double , VTotPIS_CompraNFItem double , VTotCOFINS_CompraNFItem double , " & _
                                            "VTotOutDesp_CompraNFItem double ,  VUntCustoSI_CompraNFItem double , VTotDebISSRet_CompraNFItem double , VTotIseICMS_CompraNFItem double , VTotOutICMS_CompraNFItem double , SNCredICMS_CompraNFItem double , VTotSNCredICMS_CompraNFItem double)"

''#CriacaoDeAmbiente
''#ExclusaoTabelasAuxiliares
''#CriacaoTabelasAuxiliares
''#CadastroDeParametros
''#CadastroDeTipos
''#CadastroOrigemDestino
''#EXCLUIR - USAR APENAS EM AMBIENTE DE DESENVOLVIMENTO


'' #####################################################################
'' ### #Ailton - EM TESTES
'' #####################################################################

Sub teste_CADASTRO_UNICO()
    Dim arr() As Variant: arr = Array(deleteProcessamento, Processamento): executarComandos arr
End Sub




'' #########################################################################################
'' ### #Proparts - Criacao De Ambiente para uso da nova aplicação ( Conexao NF-e e CT-e )
'' #########################################################################################

Sub main_criacao()
''==============================================================================================================='
'' OBJETIVO          : Leitura de arquivos do tipo xml (NF-e ou CT-e) para importação em banco de dados e
''                     criação de dois arquivos do tipo json dos seguintes modelos: Atualizada no ERP (lançada)
''                     e Envio do manifesto pelo ERP.
''
''==============================================================================================================='

On Error Resume Next
    Dim arr() As Variant
    
    
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    '' #EXCLUIR - USAR APENAS EM AMBIENTE DE DESENVOLVIMENTO
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    arr = Array(deleteCompras, deleteComprasItens)
    executarComandos arr
    
    
    ''#ExclusaoTabelasAuxiliares - Exclusao de tabelas auxiliares caso existam
    arr = Array(deleteProcessamento, deleteOrigemDestino, deleteParametros, deleteTipos, deleteDados)
    executarComandos arr
    
On Error GoTo 0


    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    '' #EXCLUIR - USAR APENAS EM AMBIENTE DE DESENVOLVIMENTO
    '' X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X
    arr = Array(createCompras, createComprasItens)
    executarComandos arr


    ''#CriacaoTabelasAuxiliares - Criação de tabelas auxiliares para uso no processamento de arquivos xmls e json
    arr = Array(Processamento, createOrigemDestino, createParametros, createTipos, createDados)
    executarComandos arr

    ''#CadastroDeParametros - Cadastro de parametros ex: ( Caminhos, Valores padrões e outros )
    CadastroDeItens ItensDeParametros
    
    ''#CadastroDeTipos - Cadastro de tipos para classificação de registros
    CadastroDeItens ItensDeTipos
    
    ''#CadastroOrigemDestino - Relacionamento entre campos dos arquivos (nfe,cte) das tabela (tblCompraNF,tblCompraNFItem)
    CadastroOrigemDestino
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "main_criacao"


End Sub






'' #####################################################################
'' ### #Libs - USADAS APENAS NESTE MÓDULO PARA CRIAÇÃO
'' #####################################################################

Private Sub CadastroOrigemDestino()
Dim db As Database
Dim tdf As TableDef
Dim x As Integer

    Set db = CurrentDb
    Set con = CurrentProject.Connection

    For Each Comando In carregarParametros("tblOrigemDestino", qryParametro)
        For Each tdf In db.TableDefs
           If left(tdf.Name, 4) <> "MSys" And (tdf.Name = Comando) Then
              For x = 0 To tdf.Fields.count - 1
                con.Execute Replace(Replace(qryOrigemDestino, "strDestino", tdf.Name & "." & tdf.Fields(x).Name), "strTipo", getTypeText(tdf.Fields(x).Type))
              Next x
           End If
        Next tdf
    Next Comando

Set con = Nothing

End Sub

Private Sub CadastroDeItens(Itens As Collection)
Dim con As ADODB.Connection: Set con = CurrentProject.Connection
Dim I As Variant

    For Each I In Itens
        con.Execute I
    Next I

Set con = Nothing

End Sub

Private Sub criarConsulta(nomeDaConsulta As String, scriptDaConsulta As String)
Dim db As DAO.Database: Set db = CurrentDb

    db.CreateQueryDef nomeDaConsulta, scriptDaConsulta
    db.Close

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

    '' CAMINHOS
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('caminhoDeColeta','C:\temp\proparts\Coleta\')"
    ItensDeParametros.add "INSERT INTO tblParametros (TipoDeParametro,ValorDoParametro) VALUES('caminhoDeProcessados','C:\temp\proparts\Processados\')"
    
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


