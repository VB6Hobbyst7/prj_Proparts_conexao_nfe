VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConexaoNfeCteCompraNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'
'Private Const qryProcessamento As String = "select nomecampo, valor from tblProcessamento where pk = 'strChave' and len(nomecampo)>0"
'
'
'Private Type ClassType
'        IDOLD_CompraNF As Long
'        Fil_CompraNF As String
'        NumNF_CompraNF As String
'        NumPed_CompraNF As String
'        NumOrc_CompraNF As String
'        Esp_CompraNF As String
'        Serie_CompraNF As String
'        TPNF_CompraNF As String
'        ID_NatOp_CompraNF As Long
'        ID_NatOpOLD_CompraNF As Long
'        CFOP_CompraNF As String
'        IESubsTrib_CompraNF As String
'        DTEmi_CompraNF As Date
'        DTEntd_CompraNF As Date
'        HoraEntd_CompraNF As String
'        ID_Forn_CompraNF As Long
'        ID_FornOld_CompraNF As Long
'        ID_Compr_CompraNF As Long
'        ID_Transp_CompraNF As Long
'        ID_CondPgto_CompraNF As Long
'        BaseCalcICMSSubsTrib_CompraNF As Double
'        VTotICMSSubsTrib_CompraNF As Double
'        VTotFrete_CompraNF As Double
'        VTotSeguro_CompraNF As Double
'        VTotOutDesp_CompraNF As Double
'        BaseCalcICMS_CompraNF As Double
'        VTotICMS_CompraNF As Double
'        VTotIPI_CompraNF As Double
'        VTotISS_CompraNF As Double
'        VTotProd_CompraNF As Double
'        VTotServ_CompraNF As Double
'        VTotNF_CompraNF As Double
'        TxDesc_CompraNF As Double
'        VTotDesc_CompraNF As Double
'        TPFrete_CompraNF As Double
'        Placa_CompraNF As String
'        UFVeic_CompraNF As String
'        QtdVol_CompraNF As Long
'        EspVol_CompraNF As String
'        MarcaVol_CompraNF As String
'        NumVol_CompraNF As String
'        PesoBrt_CompraNF As Double
'        PesoLiq_CompraNF As Double
'        DdAdic_CompraNF As String
'        Obs_CompraNF As String
'        Sit_CompraNF As String
'        IDCli_Depto_CompraNF As Long
'        IDCli_Contato_CompraNF As Long
'        IDCli_Email_CompraNF As Long
'        IDCli_Fone_CompraNF As Long
'        CondEsp_CompraNF As String
'        Validade_CompraNF As String
'        PzEntg_CompraNF As String
'        Garantia_CompraNF As String
'        FlagSimples_CompraNF As String
'        FlagDescBaseICMS_CompraNF As String
'        FlagExp_CompraNF As String
'        ModeloDoc_CompraNF As String
'        ChvAcesso_CompraNF As String
'        VTotPIS_CompraNF As Double
'        VTotCOFINS_CompraNF As Double
'        VTotPISRet_CompraNF As Double
'        VTotCOFINSRet_CompraNF As Double
'        VTotCSLLRet_CompraNF As Double
'        VTotIRRFRet_CompraNF As Double
'        FlagSomaST_CompraNF As String
'        FlagCalculado_CompraNF As String
'        VTotISSRet_CompraNF As Double
'        DTExt_CompraNF As Date
'        CNPJ_CPF_CompraNF As String
'        NomeCompleto_CompraNF As String
'        NomeCompleto_VendaNF As String
'        ID_Imp_CompraNF As Long
'        SitOR_CompraNF As String
'        NumOR_CompraNF As String
'        FlagNEnvWMAS_CompraNF As String
'        IDVD_CompraNF As String
'        IDVendaNF_CompraNF As String
'        FlagTransf_CompraNF As String
'End Type
'Private This As ClassType
'
'
'Property Get IDOLD_CompraNF() As Long: IDOLD_CompraNF = This.IDOLD_CompraNF: End Property
'Property Let IDOLD_CompraNF(pValue As Long): This.IDOLD_CompraNF = pValue: End Property
'Property Get Fil_CompraNF() As String: Fil_CompraNF = This.Fil_CompraNF: End Property
'Property Let Fil_CompraNF(pValue As String): This.Fil_CompraNF = pValue: End Property
'Property Get NumNF_CompraNF() As String: NumNF_CompraNF = This.NumNF_CompraNF: End Property
'Property Let NumNF_CompraNF(pValue As String): This.NumNF_CompraNF = pValue: End Property
'Property Get NumPed_CompraNF() As String: NumPed_CompraNF = This.NumPed_CompraNF: End Property
'Property Let NumPed_CompraNF(pValue As String): This.NumPed_CompraNF = pValue: End Property
'Property Get NumOrc_CompraNF() As String: NumOrc_CompraNF = This.NumOrc_CompraNF: End Property
'Property Let NumOrc_CompraNF(pValue As String): This.NumOrc_CompraNF = pValue: End Property
'Property Get Esp_CompraNF() As String: Esp_CompraNF = This.Esp_CompraNF: End Property
'Property Let Esp_CompraNF(pValue As String): This.Esp_CompraNF = pValue: End Property
'Property Get Serie_CompraNF() As String: Serie_CompraNF = This.Serie_CompraNF: End Property
'Property Let Serie_CompraNF(pValue As String): This.Serie_CompraNF = pValue: End Property
'Property Get TPNF_CompraNF() As String: TPNF_CompraNF = This.TPNF_CompraNF: End Property
'Property Let TPNF_CompraNF(pValue As String): This.TPNF_CompraNF = pValue: End Property
'Property Get ID_NatOp_CompraNF() As Long: ID_NatOp_CompraNF = This.ID_NatOp_CompraNF: End Property
'Property Let ID_NatOp_CompraNF(pValue As Long): This.ID_NatOp_CompraNF = pValue: End Property
'Property Get ID_NatOpOLD_CompraNF() As Long: ID_NatOpOLD_CompraNF = This.ID_NatOpOLD_CompraNF: End Property
'Property Let ID_NatOpOLD_CompraNF(pValue As Long): This.ID_NatOpOLD_CompraNF = pValue: End Property
'Property Get CFOP_CompraNF() As String: CFOP_CompraNF = This.CFOP_CompraNF: End Property
'Property Let CFOP_CompraNF(pValue As String): This.CFOP_CompraNF = pValue: End Property
'Property Get IESubsTrib_CompraNF() As String: IESubsTrib_CompraNF = This.IESubsTrib_CompraNF: End Property
'Property Let IESubsTrib_CompraNF(pValue As String): This.IESubsTrib_CompraNF = pValue: End Property
'Property Get DTEmi_CompraNF() As Date: DTEmi_CompraNF = This.DTEmi_CompraNF: End Property
'Property Let DTEmi_CompraNF(pValue As Date): This.DTEmi_CompraNF = pValue: End Property
'Property Get DTEntd_CompraNF() As Date: DTEntd_CompraNF = This.DTEntd_CompraNF: End Property
'Property Let DTEntd_CompraNF(pValue As Date): This.DTEntd_CompraNF = pValue: End Property
'Property Get HoraEntd_CompraNF() As String: HoraEntd_CompraNF = This.HoraEntd_CompraNF: End Property
'Property Let HoraEntd_CompraNF(pValue As String): This.HoraEntd_CompraNF = pValue: End Property
'Property Get ID_Forn_CompraNF() As Long: ID_Forn_CompraNF = This.ID_Forn_CompraNF: End Property
'Property Let ID_Forn_CompraNF(pValue As Long): This.ID_Forn_CompraNF = pValue: End Property
'Property Get ID_FornOld_CompraNF() As Long: ID_FornOld_CompraNF = This.ID_FornOld_CompraNF: End Property
'Property Let ID_FornOld_CompraNF(pValue As Long): This.ID_FornOld_CompraNF = pValue: End Property
'Property Get ID_Compr_CompraNF() As Long: ID_Compr_CompraNF = This.ID_Compr_CompraNF: End Property
'Property Let ID_Compr_CompraNF(pValue As Long): This.ID_Compr_CompraNF = pValue: End Property
'Property Get ID_Transp_CompraNF() As Long: ID_Transp_CompraNF = This.ID_Transp_CompraNF: End Property
'Property Let ID_Transp_CompraNF(pValue As Long): This.ID_Transp_CompraNF = pValue: End Property
'Property Get ID_CondPgto_CompraNF() As Long: ID_CondPgto_CompraNF = This.ID_CondPgto_CompraNF: End Property
'Property Let ID_CondPgto_CompraNF(pValue As Long): This.ID_CondPgto_CompraNF = pValue: End Property
'Property Get BaseCalcICMSSubsTrib_CompraNF() As Double: BaseCalcICMSSubsTrib_CompraNF = This.BaseCalcICMSSubsTrib_CompraNF: End Property
'Property Let BaseCalcICMSSubsTrib_CompraNF(pValue As Double): This.BaseCalcICMSSubsTrib_CompraNF = pValue: End Property
'Property Get VTotICMSSubsTrib_CompraNF() As Double: VTotICMSSubsTrib_CompraNF = This.VTotICMSSubsTrib_CompraNF: End Property
'Property Let VTotICMSSubsTrib_CompraNF(pValue As Double): This.VTotICMSSubsTrib_CompraNF = pValue: End Property
'Property Get VTotFrete_CompraNF() As Double: VTotFrete_CompraNF = This.VTotFrete_CompraNF: End Property
'Property Let VTotFrete_CompraNF(pValue As Double): This.VTotFrete_CompraNF = pValue: End Property
'Property Get VTotSeguro_CompraNF() As Double: VTotSeguro_CompraNF = This.VTotSeguro_CompraNF: End Property
'Property Let VTotSeguro_CompraNF(pValue As Double): This.VTotSeguro_CompraNF = pValue: End Property
'
'Property Get VTotOutDesp_CompraNF() As Double: VTotOutDesp_CompraNF = This.VTotOutDesp_CompraNF: End Property
'Property Let VTotOutDesp_CompraNF(pValue As Double): This.VTotOutDesp_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'
'Property Get BaseCalcICMS_CompraNF() As Double: BaseCalcICMS_CompraNF = This.BaseCalcICMS_CompraNF: End Property
'Property Let BaseCalcICMS_CompraNF(pValue As Double): This.BaseCalcICMS_CompraNF = pValue: End Property
'Property Get VTotICMS_CompraNF() As Double: VTotICMS_CompraNF = This.VTotICMS_CompraNF: End Property
'Property Let VTotICMS_CompraNF(pValue As Double): This.VTotICMS_CompraNF = pValue: End Property
'Property Get VTotIPI_CompraNF() As Double: VTotIPI_CompraNF = This.VTotIPI_CompraNF: End Property
'Property Let VTotIPI_CompraNF(pValue As Double): This.VTotIPI_CompraNF = pValue: End Property
'Property Get VTotISS_CompraNF() As Double: VTotISS_CompraNF = This.VTotISS_CompraNF: End Property
'Property Let VTotISS_CompraNF(pValue As Double): This.VTotISS_CompraNF = pValue: End Property
'Property Get VTotProd_CompraNF() As Double: VTotProd_CompraNF = This.VTotProd_CompraNF: End Property
'Property Let VTotProd_CompraNF(pValue As Double): This.VTotProd_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotServ_CompraNF() As Double: VTotServ_CompraNF = This.VTotServ_CompraNF: End Property
'Property Let VTotServ_CompraNF(pValue As Double): This.VTotServ_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotNF_CompraNF() As Double: VTotNF_CompraNF = This.VTotNF_CompraNF: End Property
'Property Let VTotNF_CompraNF(pValue As Double): This.VTotNF_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get TxDesc_CompraNF() As Double: TxDesc_CompraNF = This.TxDesc_CompraNF: End Property
'Property Let TxDesc_CompraNF(pValue As Double): This.TxDesc_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotDesc_CompraNF() As Double: VTotDesc_CompraNF = This.VTotDesc_CompraNF: End Property
'Property Let VTotDesc_CompraNF(pValue As Double): This.VTotDesc_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get TPFrete_CompraNF() As Double: TPFrete_CompraNF = This.TPFrete_CompraNF: End Property
'Property Let TPFrete_CompraNF(pValue As Double): This.TPFrete_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get Placa_CompraNF() As String: Placa_CompraNF = This.Placa_CompraNF: End Property
'Property Let Placa_CompraNF(pValue As String): This.Placa_CompraNF = pValue: End Property
'Property Get UFVeic_CompraNF() As String: UFVeic_CompraNF = This.UFVeic_CompraNF: End Property
'Property Let UFVeic_CompraNF(pValue As String): This.UFVeic_CompraNF = pValue: End Property
'Property Get QtdVol_CompraNF() As Long: QtdVol_CompraNF = This.QtdVol_CompraNF: End Property
'Property Let QtdVol_CompraNF(pValue As Long): This.QtdVol_CompraNF = pValue: End Property
'Property Get EspVol_CompraNF() As String: EspVol_CompraNF = This.EspVol_CompraNF: End Property
'Property Let EspVol_CompraNF(pValue As String): This.EspVol_CompraNF = pValue: End Property
'Property Get MarcaVol_CompraNF() As String: MarcaVol_CompraNF = This.MarcaVol_CompraNF: End Property
'Property Let MarcaVol_CompraNF(pValue As String): This.MarcaVol_CompraNF = pValue: End Property
'Property Get NumVol_CompraNF() As String: NumVol_CompraNF = This.NumVol_CompraNF: End Property
'Property Let NumVol_CompraNF(pValue As String): This.NumVol_CompraNF = pValue: End Property
'Property Get PesoBrt_CompraNF() As Double: PesoBrt_CompraNF = This.PesoBrt_CompraNF: End Property
'Property Let PesoBrt_CompraNF(pValue As Double): This.PesoBrt_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get PesoLiq_CompraNF() As Double: PesoLiq_CompraNF = This.PesoLiq_CompraNF: End Property
'Property Let PesoLiq_CompraNF(pValue As Double): This.PesoLiq_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get DdAdic_CompraNF() As String: DdAdic_CompraNF = This.DdAdic_CompraNF: End Property
'Property Let DdAdic_CompraNF(pValue As String): This.DdAdic_CompraNF = pValue: End Property
'Property Get Obs_CompraNF() As String: Obs_CompraNF = This.Obs_CompraNF: End Property
'Property Let Obs_CompraNF(pValue As String): This.Obs_CompraNF = pValue: End Property
'Property Get Sit_CompraNF() As String: Sit_CompraNF = This.Sit_CompraNF: End Property
'Property Let Sit_CompraNF(pValue As String): This.Sit_CompraNF = pValue: End Property
'Property Get IDCli_Depto_CompraNF() As Long: IDCli_Depto_CompraNF = This.IDCli_Depto_CompraNF: End Property
'Property Let IDCli_Depto_CompraNF(pValue As Long): This.IDCli_Depto_CompraNF = pValue: End Property
'Property Get IDCli_Contato_CompraNF() As Long: IDCli_Contato_CompraNF = This.IDCli_Contato_CompraNF: End Property
'Property Let IDCli_Contato_CompraNF(pValue As Long): This.IDCli_Contato_CompraNF = pValue: End Property
'Property Get IDCli_Email_CompraNF() As Long: IDCli_Email_CompraNF = This.IDCli_Email_CompraNF: End Property
'Property Let IDCli_Email_CompraNF(pValue As Long): This.IDCli_Email_CompraNF = pValue: End Property
'Property Get IDCli_Fone_CompraNF() As Long: IDCli_Fone_CompraNF = This.IDCli_Fone_CompraNF: End Property
'Property Let IDCli_Fone_CompraNF(pValue As Long): This.IDCli_Fone_CompraNF = pValue: End Property
'Property Get CondEsp_CompraNF() As String: CondEsp_CompraNF = This.CondEsp_CompraNF: End Property
'Property Let CondEsp_CompraNF(pValue As String): This.CondEsp_CompraNF = pValue: End Property
'Property Get Validade_CompraNF() As String: Validade_CompraNF = This.Validade_CompraNF: End Property
'Property Let Validade_CompraNF(pValue As String): This.Validade_CompraNF = pValue: End Property
'Property Get PzEntg_CompraNF() As String: PzEntg_CompraNF = This.PzEntg_CompraNF: End Property
'Property Let PzEntg_CompraNF(pValue As String): This.PzEntg_CompraNF = pValue: End Property
'Property Get Garantia_CompraNF() As String: Garantia_CompraNF = This.Garantia_CompraNF: End Property
'Property Let Garantia_CompraNF(pValue As String): This.Garantia_CompraNF = pValue: End Property
'Property Get FlagSimples_CompraNF() As String: FlagSimples_CompraNF = This.FlagSimples_CompraNF: End Property
'Property Let FlagSimples_CompraNF(pValue As String): This.FlagSimples_CompraNF = pValue: End Property
'Property Get FlagDescBaseICMS_CompraNF() As String: FlagDescBaseICMS_CompraNF = This.FlagDescBaseICMS_CompraNF: End Property
'Property Let FlagDescBaseICMS_CompraNF(pValue As String): This.FlagDescBaseICMS_CompraNF = pValue: End Property
'Property Get FlagExp_CompraNF() As String: FlagExp_CompraNF = This.FlagExp_CompraNF: End Property
'Property Let FlagExp_CompraNF(pValue As String): This.FlagExp_CompraNF = pValue: End Property
'Property Get ModeloDoc_CompraNF() As String: ModeloDoc_CompraNF = This.ModeloDoc_CompraNF: End Property
'Property Let ModeloDoc_CompraNF(pValue As String): This.ModeloDoc_CompraNF = pValue: End Property
'Property Get ChvAcesso_CompraNF() As String: ChvAcesso_CompraNF = This.ChvAcesso_CompraNF: End Property
'Property Let ChvAcesso_CompraNF(pValue As String): This.ChvAcesso_CompraNF = pValue: End Property
'Property Get VTotPIS_CompraNF() As Double: VTotPIS_CompraNF = This.VTotPIS_CompraNF: End Property
'Property Let VTotPIS_CompraNF(pValue As Double): This.VTotPIS_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotCOFINS_CompraNF() As Double: VTotCOFINS_CompraNF = This.VTotCOFINS_CompraNF: End Property
'Property Let VTotCOFINS_CompraNF(pValue As Double): This.VTotCOFINS_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotPISRet_CompraNF() As Double: VTotPISRet_CompraNF = This.VTotPISRet_CompraNF: End Property
'Property Let VTotPISRet_CompraNF(pValue As Double): This.VTotPISRet_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotCOFINSRet_CompraNF() As Double: VTotCOFINSRet_CompraNF = This.VTotCOFINSRet_CompraNF: End Property
'Property Let VTotCOFINSRet_CompraNF(pValue As Double): This.VTotCOFINSRet_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotCSLLRet_CompraNF() As Double: VTotCSLLRet_CompraNF = This.VTotCSLLRet_CompraNF: End Property
'Property Let VTotCSLLRet_CompraNF(pValue As Double): This.VTotCSLLRet_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get VTotIRRFRet_CompraNF() As Double: VTotIRRFRet_CompraNF = This.VTotIRRFRet_CompraNF: End Property
'Property Let VTotIRRFRet_CompraNF(pValue As Double): This.VTotIRRFRet_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get FlagSomaST_CompraNF() As String: FlagSomaST_CompraNF = This.FlagSomaST_CompraNF: End Property
'Property Let FlagSomaST_CompraNF(pValue As String): This.FlagSomaST_CompraNF = pValue: End Property
'Property Get FlagCalculado_CompraNF() As String: FlagCalculado_CompraNF = This.FlagCalculado_CompraNF: End Property
'Property Let FlagCalculado_CompraNF(pValue As String): This.FlagCalculado_CompraNF = pValue: End Property
'Property Get VTotISSRet_CompraNF() As Double: VTotISSRet_CompraNF = This.VTotISSRet_CompraNF: End Property
'Property Let VTotISSRet_CompraNF(pValue As Double): This.VTotISSRet_CompraNF = Val(Replace(pValue, ".", ",")): End Property
'Property Get DTExt_CompraNF() As Date: DTExt_CompraNF = This.DTExt_CompraNF: End Property
'Property Let DTExt_CompraNF(pValue As Date): This.DTExt_CompraNF = pValue: End Property
'Property Get CNPJ_CPF_CompraNF() As String: CNPJ_CPF_CompraNF = This.CNPJ_CPF_CompraNF: End Property
'Property Let CNPJ_CPF_CompraNF(pValue As String): This.CNPJ_CPF_CompraNF = pValue: End Property
'Property Get NomeCompleto_CompraNF() As String: NomeCompleto_CompraNF = This.NomeCompleto_CompraNF: End Property
'Property Let NomeCompleto_CompraNF(pValue As String): This.NomeCompleto_CompraNF = pValue: End Property
'Property Get NomeCompleto_VendaNF() As String: NomeCompleto_VendaNF = This.NomeCompleto_VendaNF: End Property
'Property Let NomeCompleto_VendaNF(pValue As String): This.NomeCompleto_VendaNF = pValue: End Property
'Property Get ID_Imp_CompraNF() As Long: ID_Imp_CompraNF = This.ID_Imp_CompraNF: End Property
'Property Let ID_Imp_CompraNF(pValue As Long): This.ID_Imp_CompraNF = pValue: End Property
'Property Get SitOR_CompraNF() As String: SitOR_CompraNF = This.SitOR_CompraNF: End Property
'Property Let SitOR_CompraNF(pValue As String): This.SitOR_CompraNF = pValue: End Property
'Property Get NumOR_CompraNF() As String: NumOR_CompraNF = This.NumOR_CompraNF: End Property
'Property Let NumOR_CompraNF(pValue As String): This.NumOR_CompraNF = pValue: End Property
'Property Get FlagNEnvWMAS_CompraNF() As String: FlagNEnvWMAS_CompraNF = This.FlagNEnvWMAS_CompraNF: End Property
'Property Let FlagNEnvWMAS_CompraNF(pValue As String): This.FlagNEnvWMAS_CompraNF = pValue: End Property
'Property Get IDVD_CompraNF() As String: IDVD_CompraNF = This.IDVD_CompraNF: End Property
'Property Let IDVD_CompraNF(pValue As String): This.IDVD_CompraNF = pValue: End Property
'Property Get IDVendaNF_CompraNF() As String: IDVendaNF_CompraNF = This.IDVendaNF_CompraNF: End Property
'Property Let IDVendaNF_CompraNF(pValue As String): This.IDVendaNF_CompraNF = pValue: End Property
'Property Get FlagTransf_CompraNF() As String: FlagTransf_CompraNF = This.FlagTransf_CompraNF: End Property
'Property Let FlagTransf_CompraNF(pValue As String): This.FlagTransf_CompraNF = pValue: End Property
'
'
'Sub carregar_dados(strChave As String)
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(Replace(qryProcessamento, "strChave", strChave))
'
'
'Do While Not rst.EOF
'        Debug.Print rst.Fields("nomecampo").Value
'        Debug.Print rst.Fields("valor").Value
'
'        Select Case rst.Fields("nomecampo").Value
'
'
'                '' #AILTON - qryUpdate
''                Case "Fil_CompraNF": Fil_CompraNF = rst.Fields("valor").Value
''                Case "ID_NatOp_CompraNF": ID_NatOp_CompraNF = rst.Fields("valor").Value
''                Case "CFOP_CompraNF": CFOP_CompraNF = rst.Fields("valor").Value
''                Case "NumPed_CompraNF": NumPed_CompraNF = rst.Fields("valor").Value
'
'                Case "NumNF_CompraNF": NumNF_CompraNF = rst.Fields("valor").Value
'                Case "ModeloDoc_CompraNF": ModeloDoc_CompraNF = rst.Fields("valor").Value
'                Case "Serie_CompraNF": Serie_CompraNF = rst.Fields("valor").Value
'                Case "TPNF_CompraNF": TPNF_CompraNF = rst.Fields("valor").Value
'
'                '' #AILTON - qryTransform
'                Case "DTEmi_CompraNF": DTEmi_CompraNF = CDate(Replace(Mid(rst.Fields("valor").Value, 1, 10), "-", "/"))
'                Case "DTEntd_CompraNF": DTEntd_CompraNF = Now()
'
'                Case "ID_Forn_CompraNF": ID_Forn_CompraNF = rst.Fields("valor").Value
'                Case "BaseCalcICMS_CompraNF": BaseCalcICMS_CompraNF = rst.Fields("valor").Value
'                Case "Sit_CompraNF": Sit_CompraNF = rst.Fields("valor").Value
'                Case "ChvAcesso_CompraNF": ChvAcesso_CompraNF = rst.Fields("valor").Value
'
'
'                '' #AILTON - CORRECAO DE FORMATO
'                Case "BaseCalcICMSSubsTrib_CompraNF": BaseCalcICMSSubsTrib_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotICMS_CompraNF": VTotICMS_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotServ_CompraNF": VTotServ_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotProd_CompraNF": VTotProd_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotNF_CompraNF": VTotNF_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotICMSSubsTrib_CompraNF": VTotICMSSubsTrib_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotFrete_CompraNF": VTotFrete_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotSeguro_CompraNF": VTotSeguro_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotOutDesp_CompraNF": VTotOutDesp_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotIPI_CompraNF": VTotIPI_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotISS_CompraNF": VTotISS_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "TxDesc_CompraNF": TxDesc_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'                Case "VTotDesc_CompraNF": VTotDesc_CompraNF = Format(Replace(rst.Fields("valor").Value, ".", ","), "#,##0.00")
'
'                Case "IDVD_CompraNF": IDVD_CompraNF = rst.Fields("valor").Value
'                Case "CNPJ_CPF": CNPJ_CPF_CompraNF = Replace(Replace(Replace(rst.Fields("valor").Value, ".", ""), "-", ""), "/", "")
'
'        End Select
'        rst.MoveNext
'Loop
'
'
'db.Close
'Set db = Nothing
'End Sub
'
'
'Sub cadastrar()
'Dim db As DAO.Database: Set db = CurrentDb
'Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblCompraNF")
'Dim parts() As String
'
'
'With qdf
'
'        '' #AILTON
'        '' qrySelectEmpresa_FiltroFil - SELECT tblEmpresa.ID_Empresa, tblEmpresa.RzSoc_Empresa, tblEmpresa.CNPJ_Empresa FROM tblEmpresa WHERE (((tblEmpresa.CNPJ_Empresa) Is Not Null));
'
'        'Replace(Replace(Replace([CNPJ_CPF],""."",""""),""-"",""""),""/"",""""))
'
'        If Fil_CompraNF = "" Then
'            .Parameters("Fil_CompraNF") = 0
'        Else
'            .Parameters("Fil_CompraNF") = Fil_CompraNF
'        End If
'
'        .Parameters("CNPJ_CPF_CompraNF") = CNPJ_CPF_CompraNF
'
''        .Parameters("IDOLD_CompraNF") = IDOLD_CompraNF
''        .Parameters("CFOP_CompraNF") = CFOP_CompraNF
''        .Parameters("NumPed_CompraNF") = NumPed_CompraNF
'
'
'        .Parameters("NumNF_CompraNF") = NumNF_CompraNF
'        .Parameters("ModeloDoc_CompraNF") = ModeloDoc_CompraNF
'        .Parameters("Serie_CompraNF") = Serie_CompraNF
'        .Parameters("TPNF_CompraNF") = TPNF_CompraNF
'
'
'        .Parameters("DTEmi_CompraNF") = DTEmi_CompraNF
'        .Parameters("DTEntd_CompraNF") = Date 'DTEntd_CompraNF
'
'        .Parameters("ID_Forn_CompraNF") = ID_Forn_CompraNF
'        .Parameters("BaseCalcICMS_CompraNF") = BaseCalcICMS_CompraNF
'        .Parameters("VTotICMS_CompraNF") = VTotICMS_CompraNF
'        .Parameters("VTotServ_CompraNF") = VTotServ_CompraNF
'
'        '' #AILTON - R$ 69.249,00 - 69249.39
'        .Parameters("VTotProd_CompraNF") = VTotProd_CompraNF
'        .Parameters("VTotNF_CompraNF") = VTotNF_CompraNF
'        .Parameters("Sit_CompraNF") = Sit_CompraNF
'        .Parameters("ChvAcesso_CompraNF") = ChvAcesso_CompraNF
'        .Parameters("BaseCalcICMSSubsTrib_CompraNF") = BaseCalcICMSSubsTrib_CompraNF
'        .Parameters("VTotICMSSubsTrib_CompraNF") = VTotICMSSubsTrib_CompraNF
'        .Parameters("VTotFrete_CompraNF") = VTotFrete_CompraNF
'        .Parameters("VTotSeguro_CompraNF") = VTotSeguro_CompraNF
'        .Parameters("VTotOutDesp_CompraNF") = VTotOutDesp_CompraNF
'        .Parameters("VTotIPI_CompraNF") = VTotIPI_CompraNF
'        .Parameters("VTotISS_CompraNF") = VTotISS_CompraNF
'        .Parameters("TxDesc_CompraNF") = TxDesc_CompraNF
'        .Parameters("VTotDesc_CompraNF") = VTotDesc_CompraNF
'
'        If Fil_CompraNF = "4" Then
'            .Parameters("IDVD_CompraNF") = Replace(parts(LBound(Split((IDVD_CompraNF), ","))), "Pedido:", "")
'        Else
'            .Parameters("IDVD_CompraNF") = 0
'        End If
'
'        '' #URGENTE
'        .Execute
'End With
'End Sub
'
