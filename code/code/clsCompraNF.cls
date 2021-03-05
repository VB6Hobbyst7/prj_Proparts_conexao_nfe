VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompraNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private Type ClassType
    ID_CompraNF As String
    IDOLD_CompraNF As String
    Fil_CompraNF As String
    NumNF_CompraNF As String
    NumPed_CompraNF As String
    NumOrc_CompraNF As String
    Esp_CompraNF As String
    Serie_CompraNF As String
    TPNF_CompraNF As String
    ID_NatOp_CompraNF As String
    ID_NatOpOLD_CompraNF As String
    CFOP_CompraNF As String
    IESubsTrib_CompraNF As String
    DTEmi_CompraNF As String
    DTEntd_CompraNF As String
    HoraEntd_CompraNF As String
    ID_Forn_CompraNF As String
    ID_FornOld_CompraNF As String
    ID_Compr_CompraNF As String
    ID_Transp_CompraNF As String
    ID_CondPgto_CompraNF As String
    BaseCalcICMSSubsTrib_CompraNF As String
    VTotICMSSubsTrib_CompraNF As String
    VTotFrete_CompraNF As String
    VTotSeguro_CompraNF As String
    VTotOutDesp_CompraNF As String
    BaseCalcICMS_CompraNF As String
    VTotICMS_CompraNF As String
    VTotIPI_CompraNF As String
    VTotISS_CompraNF As String
    VTotProd_CompraNF As String
    VTotServ_CompraNF As String
    VTotNF_CompraNF As String
    TxDesc_CompraNF As String
    VTotDesc_CompraNF As String
    TPFrete_CompraNF As String
    Placa_CompraNF As String
    UFVeic_CompraNF As String
    QtdVol_CompraNF As String
    EspVol_CompraNF As String
    MarcaVol_CompraNF As String
    NumVol_CompraNF As String
    PesoBrt_CompraNF As String
    PesoLiq_CompraNF As String
    DdAdic_CompraNF As String
    Obs_CompraNF As String
    Sit_CompraNF As String
    IDCli_Depto_CompraNF As String
    IDCli_Contato_CompraNF As String
    IDCli_Email_CompraNF As String
    IDCli_Fone_CompraNF As String
    CondEsp_CompraNF As String
    Validade_CompraNF As String
    PzEntg_CompraNF As String
    Garantia_CompraNF As String
    FlagSimples_CompraNF As String
    FlagDescBaseICMS_CompraNF As String
    FlagExp_CompraNF As String
    ModeloDoc_CompraNF As String
    ChvAcesso_CompraNF As String
    VTotPIS_CompraNF As String
    VTotCOFINS_CompraNF As String
    VTotPISRet_CompraNF As String
    VTotCOFINSRet_CompraNF As String
    VTotCSLLRet_CompraNF As String
    VTotIRRFRet_CompraNF As String
    FlagSomaST_CompraNF As String
    FlagCalculado_CompraNF As String
    VTotISSRet_CompraNF As String
    DTExt_CompraNF As String
    CNPJ_CPF_CompraNF As String
    NomeCompleto_CompraNF As String
    NomeCompleto_VendaNF As String
    ID_Imp_CompraNF As String
    SitOR_CompraNF As String
    NumOR_CompraNF As String
    FlagNEnvWMAS_CompraNF As String
    IDVD_CompraNF As String
    IDVendaNF_CompraNF As String
    FlagTransf_CompraNF As String
    CompraNFItem As New Collection
End Type
Private This As ClassType


'Public Function ProcessarArquivo(pFolderSource As String)
'
'Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
'Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
'
'
'    For Each fileName In GetFilesInFolder(pFolderSource)
'        If (InStrRev(fileName, ".xml", -1, vbTextCompare)) Then
'            XDoc.Load fileName
'
'            ID_CompraNF = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("mod")(0).Text
'            ID_CompraNF = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("mod")(0).Text
'            ID_CompraNF = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("mod")(0).Text
'            ID_CompraNF = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("mod")(0).Text
'            ID_CompraNF = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("mod")(0).Text
'
'        End If
'
'    Next fileName
'
'Set XDoc = Nothing
'
'
'
'
'End Function



'Public Function adicionarItens(pItem As clsCompraNFItem)
'
'CompraNFItem.Add pItem
'
'End Function

'Property Get CompraNFItem() As Collection
'
'CompraNFItem = This.CompraNFItem
'
'End Property

Property Get ID_CompraNF() As String ' formatação

    ID_CompraNF = This.ID_CompraNF
    
End Property



Property Let ID_CompraNF(pValue As String):             This.ID_CompraNF = pValue:              End Property
Property Get IDOLD_CompraNF() As String:                   IDOLD_CompraNF = This.IDOLD_CompraNF:          End Property
Property Let IDOLD_CompraNF(pValue As String):             This.IDOLD_CompraNF = pValue:              End Property
Property Get Fil_CompraNF() As String:                   Fil_CompraNF = This.Fil_CompraNF:          End Property
Property Let Fil_CompraNF(pValue As String):             This.Fil_CompraNF = pValue:              End Property
Property Get NumNF_CompraNF() As String:                   NumNF_CompraNF = This.NumNF_CompraNF:          End Property
Property Let NumNF_CompraNF(pValue As String):             This.NumNF_CompraNF = pValue:              End Property
Property Get NumPed_CompraNF() As String:                   NumPed_CompraNF = This.NumPed_CompraNF:          End Property
Property Let NumPed_CompraNF(pValue As String):             This.NumPed_CompraNF = pValue:              End Property
Property Get NumOrc_CompraNF() As String:                   NumOrc_CompraNF = This.NumOrc_CompraNF:          End Property
Property Let NumOrc_CompraNF(pValue As String):             This.NumOrc_CompraNF = pValue:              End Property
Property Get Esp_CompraNF() As String:                   Esp_CompraNF = This.Esp_CompraNF:          End Property
Property Let Esp_CompraNF(pValue As String):             This.Esp_CompraNF = pValue:              End Property
Property Get Serie_CompraNF() As String:                   Serie_CompraNF = This.Serie_CompraNF:          End Property
Property Let Serie_CompraNF(pValue As String):             This.Serie_CompraNF = pValue:              End Property
Property Get TPNF_CompraNF() As String:                   TPNF_CompraNF = This.TPNF_CompraNF:          End Property
Property Let TPNF_CompraNF(pValue As String):             This.TPNF_CompraNF = pValue:              End Property
Property Get ID_NatOp_CompraNF() As String:                   ID_NatOp_CompraNF = This.ID_NatOp_CompraNF:          End Property
Property Let ID_NatOp_CompraNF(pValue As String):             This.ID_NatOp_CompraNF = pValue:              End Property
Property Get ID_NatOpOLD_CompraNF() As String:                   ID_NatOpOLD_CompraNF = This.ID_NatOpOLD_CompraNF:          End Property
Property Let ID_NatOpOLD_CompraNF(pValue As String):             This.ID_NatOpOLD_CompraNF = pValue:              End Property
Property Get CFOP_CompraNF() As String:                   CFOP_CompraNF = This.CFOP_CompraNF:          End Property
Property Let CFOP_CompraNF(pValue As String):             This.CFOP_CompraNF = pValue:              End Property
Property Get IESubsTrib_CompraNF() As String:                   IESubsTrib_CompraNF = This.IESubsTrib_CompraNF:          End Property
Property Let IESubsTrib_CompraNF(pValue As String):             This.IESubsTrib_CompraNF = pValue:              End Property
Property Get DTEmi_CompraNF() As String:                   DTEmi_CompraNF = This.DTEmi_CompraNF:          End Property
Property Let DTEmi_CompraNF(pValue As String):             This.DTEmi_CompraNF = pValue:              End Property
Property Get DTEntd_CompraNF() As String:                   DTEntd_CompraNF = This.DTEntd_CompraNF:          End Property
Property Let DTEntd_CompraNF(pValue As String):             This.DTEntd_CompraNF = pValue:              End Property
Property Get HoraEntd_CompraNF() As String:                   HoraEntd_CompraNF = This.HoraEntd_CompraNF:          End Property
Property Let HoraEntd_CompraNF(pValue As String):             This.HoraEntd_CompraNF = pValue:              End Property
Property Get ID_Forn_CompraNF() As String:                   ID_Forn_CompraNF = This.ID_Forn_CompraNF:          End Property
Property Let ID_Forn_CompraNF(pValue As String):             This.ID_Forn_CompraNF = pValue:              End Property
Property Get ID_FornOld_CompraNF() As String:                   ID_FornOld_CompraNF = This.ID_FornOld_CompraNF:          End Property
Property Let ID_FornOld_CompraNF(pValue As String):             This.ID_FornOld_CompraNF = pValue:              End Property
Property Get ID_Compr_CompraNF() As String:                   ID_Compr_CompraNF = This.ID_Compr_CompraNF:          End Property
Property Let ID_Compr_CompraNF(pValue As String):             This.ID_Compr_CompraNF = pValue:              End Property
Property Get ID_Transp_CompraNF() As String:                   ID_Transp_CompraNF = This.ID_Transp_CompraNF:          End Property
Property Let ID_Transp_CompraNF(pValue As String):             This.ID_Transp_CompraNF = pValue:              End Property
Property Get ID_CondPgto_CompraNF() As String:                   ID_CondPgto_CompraNF = This.ID_CondPgto_CompraNF:          End Property
Property Let ID_CondPgto_CompraNF(pValue As String):             This.ID_CondPgto_CompraNF = pValue:              End Property
Property Get BaseCalcICMSSubsTrib_CompraNF() As String:                   BaseCalcICMSSubsTrib_CompraNF = This.BaseCalcICMSSubsTrib_CompraNF:          End Property
Property Let BaseCalcICMSSubsTrib_CompraNF(pValue As String):             This.BaseCalcICMSSubsTrib_CompraNF = pValue:              End Property
Property Get VTotICMSSubsTrib_CompraNF() As String:                   VTotICMSSubsTrib_CompraNF = This.VTotICMSSubsTrib_CompraNF:          End Property
Property Let VTotICMSSubsTrib_CompraNF(pValue As String):             This.VTotICMSSubsTrib_CompraNF = pValue:              End Property
Property Get VTotFrete_CompraNF() As String:                   VTotFrete_CompraNF = This.VTotFrete_CompraNF:          End Property
Property Let VTotFrete_CompraNF(pValue As String):             This.VTotFrete_CompraNF = pValue:              End Property
Property Get VTotSeguro_CompraNF() As String:                   VTotSeguro_CompraNF = This.VTotSeguro_CompraNF:          End Property
Property Let VTotSeguro_CompraNF(pValue As String):             This.VTotSeguro_CompraNF = pValue:              End Property
Property Get VTotOutDesp_CompraNF() As String:                   VTotOutDesp_CompraNF = This.VTotOutDesp_CompraNF:          End Property
Property Let VTotOutDesp_CompraNF(pValue As String):             This.VTotOutDesp_CompraNF = pValue:              End Property
Property Get BaseCalcICMS_CompraNF() As String:                   BaseCalcICMS_CompraNF = This.BaseCalcICMS_CompraNF:          End Property
Property Let BaseCalcICMS_CompraNF(pValue As String):             This.BaseCalcICMS_CompraNF = pValue:              End Property
Property Get VTotICMS_CompraNF() As String:                   VTotICMS_CompraNF = This.VTotICMS_CompraNF:          End Property
Property Let VTotICMS_CompraNF(pValue As String):             This.VTotICMS_CompraNF = pValue:              End Property
Property Get VTotIPI_CompraNF() As String:                   VTotIPI_CompraNF = This.VTotIPI_CompraNF:          End Property
Property Let VTotIPI_CompraNF(pValue As String):             This.VTotIPI_CompraNF = pValue:              End Property
Property Get VTotISS_CompraNF() As String:                   VTotISS_CompraNF = This.VTotISS_CompraNF:          End Property
Property Let VTotISS_CompraNF(pValue As String):             This.VTotISS_CompraNF = pValue:              End Property
Property Get VTotProd_CompraNF() As String:                   VTotProd_CompraNF = This.VTotProd_CompraNF:          End Property
Property Let VTotProd_CompraNF(pValue As String):             This.VTotProd_CompraNF = pValue:              End Property
Property Get VTotServ_CompraNF() As String:                   VTotServ_CompraNF = This.VTotServ_CompraNF:          End Property
Property Let VTotServ_CompraNF(pValue As String):             This.VTotServ_CompraNF = pValue:              End Property
Property Get VTotNF_CompraNF() As String:                   VTotNF_CompraNF = This.VTotNF_CompraNF:          End Property
Property Let VTotNF_CompraNF(pValue As String):             This.VTotNF_CompraNF = pValue:              End Property
Property Get TxDesc_CompraNF() As String:                   TxDesc_CompraNF = This.TxDesc_CompraNF:          End Property
Property Let TxDesc_CompraNF(pValue As String):             This.TxDesc_CompraNF = pValue:              End Property
Property Get VTotDesc_CompraNF() As String:                   VTotDesc_CompraNF = This.VTotDesc_CompraNF:          End Property
Property Let VTotDesc_CompraNF(pValue As String):             This.VTotDesc_CompraNF = pValue:              End Property
Property Get TPFrete_CompraNF() As String:                   TPFrete_CompraNF = This.TPFrete_CompraNF:          End Property
Property Let TPFrete_CompraNF(pValue As String):             This.TPFrete_CompraNF = pValue:              End Property
Property Get Placa_CompraNF() As String:                   Placa_CompraNF = This.Placa_CompraNF:          End Property
Property Let Placa_CompraNF(pValue As String):             This.Placa_CompraNF = pValue:              End Property
Property Get UFVeic_CompraNF() As String:                   UFVeic_CompraNF = This.UFVeic_CompraNF:          End Property
Property Let UFVeic_CompraNF(pValue As String):             This.UFVeic_CompraNF = pValue:              End Property
Property Get QtdVol_CompraNF() As String:                   QtdVol_CompraNF = This.QtdVol_CompraNF:          End Property
Property Let QtdVol_CompraNF(pValue As String):             This.QtdVol_CompraNF = pValue:              End Property
Property Get EspVol_CompraNF() As String:                   EspVol_CompraNF = This.EspVol_CompraNF:          End Property
Property Let EspVol_CompraNF(pValue As String):             This.EspVol_CompraNF = pValue:              End Property
Property Get MarcaVol_CompraNF() As String:                   MarcaVol_CompraNF = This.MarcaVol_CompraNF:          End Property
Property Let MarcaVol_CompraNF(pValue As String):             This.MarcaVol_CompraNF = pValue:              End Property
Property Get NumVol_CompraNF() As String:                   NumVol_CompraNF = This.NumVol_CompraNF:          End Property
Property Let NumVol_CompraNF(pValue As String):             This.NumVol_CompraNF = pValue:              End Property
Property Get PesoBrt_CompraNF() As String:                   PesoBrt_CompraNF = This.PesoBrt_CompraNF:          End Property
Property Let PesoBrt_CompraNF(pValue As String):             This.PesoBrt_CompraNF = pValue:              End Property
Property Get PesoLiq_CompraNF() As String:                   PesoLiq_CompraNF = This.PesoLiq_CompraNF:          End Property
Property Let PesoLiq_CompraNF(pValue As String):             This.PesoLiq_CompraNF = pValue:              End Property
Property Get DdAdic_CompraNF() As String:                   DdAdic_CompraNF = This.DdAdic_CompraNF:          End Property
Property Let DdAdic_CompraNF(pValue As String):             This.DdAdic_CompraNF = pValue:              End Property
Property Get Obs_CompraNF() As String:                   Obs_CompraNF = This.Obs_CompraNF:          End Property
Property Let Obs_CompraNF(pValue As String):             This.Obs_CompraNF = pValue:              End Property
Property Get Sit_CompraNF() As String:                   Sit_CompraNF = This.Sit_CompraNF:          End Property
Property Let Sit_CompraNF(pValue As String):             This.Sit_CompraNF = pValue:              End Property
Property Get IDCli_Depto_CompraNF() As String:                   IDCli_Depto_CompraNF = This.IDCli_Depto_CompraNF:          End Property
Property Let IDCli_Depto_CompraNF(pValue As String):             This.IDCli_Depto_CompraNF = pValue:              End Property
Property Get IDCli_Contato_CompraNF() As String:                   IDCli_Contato_CompraNF = This.IDCli_Contato_CompraNF:          End Property
Property Let IDCli_Contato_CompraNF(pValue As String):             This.IDCli_Contato_CompraNF = pValue:              End Property
Property Get IDCli_Email_CompraNF() As String:                   IDCli_Email_CompraNF = This.IDCli_Email_CompraNF:          End Property
Property Let IDCli_Email_CompraNF(pValue As String):             This.IDCli_Email_CompraNF = pValue:              End Property
Property Get IDCli_Fone_CompraNF() As String:                   IDCli_Fone_CompraNF = This.IDCli_Fone_CompraNF:          End Property
Property Let IDCli_Fone_CompraNF(pValue As String):             This.IDCli_Fone_CompraNF = pValue:              End Property
Property Get CondEsp_CompraNF() As String:                   CondEsp_CompraNF = This.CondEsp_CompraNF:          End Property
Property Let CondEsp_CompraNF(pValue As String):             This.CondEsp_CompraNF = pValue:              End Property
Property Get Validade_CompraNF() As String:                   Validade_CompraNF = This.Validade_CompraNF:          End Property
Property Let Validade_CompraNF(pValue As String):             This.Validade_CompraNF = pValue:              End Property
Property Get PzEntg_CompraNF() As String:                   PzEntg_CompraNF = This.PzEntg_CompraNF:          End Property
Property Let PzEntg_CompraNF(pValue As String):             This.PzEntg_CompraNF = pValue:              End Property
Property Get Garantia_CompraNF() As String:                   Garantia_CompraNF = This.Garantia_CompraNF:          End Property
Property Let Garantia_CompraNF(pValue As String):             This.Garantia_CompraNF = pValue:              End Property
Property Get FlagSimples_CompraNF() As String:                   FlagSimples_CompraNF = This.FlagSimples_CompraNF:          End Property
Property Let FlagSimples_CompraNF(pValue As String):             This.FlagSimples_CompraNF = pValue:              End Property
Property Get FlagDescBaseICMS_CompraNF() As String:                   FlagDescBaseICMS_CompraNF = This.FlagDescBaseICMS_CompraNF:          End Property
Property Let FlagDescBaseICMS_CompraNF(pValue As String):             This.FlagDescBaseICMS_CompraNF = pValue:              End Property
Property Get FlagExp_CompraNF() As String:                   FlagExp_CompraNF = This.FlagExp_CompraNF:          End Property
Property Let FlagExp_CompraNF(pValue As String):             This.FlagExp_CompraNF = pValue:              End Property
Property Get ModeloDoc_CompraNF() As String:                   ModeloDoc_CompraNF = This.ModeloDoc_CompraNF:          End Property
Property Let ModeloDoc_CompraNF(pValue As String):             This.ModeloDoc_CompraNF = pValue:              End Property
Property Get ChvAcesso_CompraNF() As String:                   ChvAcesso_CompraNF = This.ChvAcesso_CompraNF:          End Property
Property Let ChvAcesso_CompraNF(pValue As String):             This.ChvAcesso_CompraNF = pValue:              End Property
Property Get VTotPIS_CompraNF() As String:                   VTotPIS_CompraNF = This.VTotPIS_CompraNF:          End Property
Property Let VTotPIS_CompraNF(pValue As String):             This.VTotPIS_CompraNF = pValue:              End Property
Property Get VTotCOFINS_CompraNF() As String:                   VTotCOFINS_CompraNF = This.VTotCOFINS_CompraNF:          End Property
Property Let VTotCOFINS_CompraNF(pValue As String):             This.VTotCOFINS_CompraNF = pValue:              End Property
Property Get VTotPISRet_CompraNF() As String:                   VTotPISRet_CompraNF = This.VTotPISRet_CompraNF:          End Property
Property Let VTotPISRet_CompraNF(pValue As String):             This.VTotPISRet_CompraNF = pValue:              End Property
Property Get VTotCOFINSRet_CompraNF() As String:                   VTotCOFINSRet_CompraNF = This.VTotCOFINSRet_CompraNF:          End Property
Property Let VTotCOFINSRet_CompraNF(pValue As String):             This.VTotCOFINSRet_CompraNF = pValue:              End Property
Property Get VTotCSLLRet_CompraNF() As String:                   VTotCSLLRet_CompraNF = This.VTotCSLLRet_CompraNF:          End Property
Property Let VTotCSLLRet_CompraNF(pValue As String):             This.VTotCSLLRet_CompraNF = pValue:              End Property
Property Get VTotIRRFRet_CompraNF() As String:                   VTotIRRFRet_CompraNF = This.VTotIRRFRet_CompraNF:          End Property
Property Let VTotIRRFRet_CompraNF(pValue As String):             This.VTotIRRFRet_CompraNF = pValue:              End Property
Property Get FlagSomaST_CompraNF() As String:                   FlagSomaST_CompraNF = This.FlagSomaST_CompraNF:          End Property
Property Let FlagSomaST_CompraNF(pValue As String):             This.FlagSomaST_CompraNF = pValue:              End Property
Property Get FlagCalculado_CompraNF() As String:                   FlagCalculado_CompraNF = This.FlagCalculado_CompraNF:          End Property
Property Let FlagCalculado_CompraNF(pValue As String):             This.FlagCalculado_CompraNF = pValue:              End Property
Property Get VTotISSRet_CompraNF() As String:                   VTotISSRet_CompraNF = This.VTotISSRet_CompraNF:          End Property
Property Let VTotISSRet_CompraNF(pValue As String):             This.VTotISSRet_CompraNF = pValue:              End Property
Property Get DTExt_CompraNF() As String:                   DTExt_CompraNF = This.DTExt_CompraNF:          End Property
Property Let DTExt_CompraNF(pValue As String):             This.DTExt_CompraNF = pValue:              End Property
Property Get CNPJ_CPF_CompraNF() As String:                   CNPJ_CPF_CompraNF = This.CNPJ_CPF_CompraNF:          End Property
Property Let CNPJ_CPF_CompraNF(pValue As String):             This.CNPJ_CPF_CompraNF = pValue:              End Property
Property Get NomeCompleto_CompraNF() As String:                   NomeCompleto_CompraNF = This.NomeCompleto_CompraNF:          End Property
Property Let NomeCompleto_CompraNF(pValue As String):             This.NomeCompleto_CompraNF = pValue:              End Property
Property Get NomeCompleto_VendaNF() As String:                   NomeCompleto_VendaNF = This.NomeCompleto_VendaNF:          End Property
Property Let NomeCompleto_VendaNF(pValue As String):             This.NomeCompleto_VendaNF = pValue:              End Property
Property Get ID_Imp_CompraNF() As String:                   ID_Imp_CompraNF = This.ID_Imp_CompraNF:          End Property
Property Let ID_Imp_CompraNF(pValue As String):             This.ID_Imp_CompraNF = pValue:              End Property
Property Get SitOR_CompraNF() As String:                   SitOR_CompraNF = This.SitOR_CompraNF:          End Property
Property Let SitOR_CompraNF(pValue As String):             This.SitOR_CompraNF = pValue:              End Property
Property Get NumOR_CompraNF() As String:                   NumOR_CompraNF = This.NumOR_CompraNF:          End Property
Property Let NumOR_CompraNF(pValue As String):             This.NumOR_CompraNF = pValue:              End Property
Property Get FlagNEnvWMAS_CompraNF() As String:                   FlagNEnvWMAS_CompraNF = This.FlagNEnvWMAS_CompraNF:          End Property
Property Let FlagNEnvWMAS_CompraNF(pValue As String):             This.FlagNEnvWMAS_CompraNF = pValue:              End Property
Property Get IDVD_CompraNF() As String:                   IDVD_CompraNF = This.IDVD_CompraNF:          End Property
Property Let IDVD_CompraNF(pValue As String):             This.IDVD_CompraNF = pValue:              End Property
Property Get IDVendaNF_CompraNF() As String:                   IDVendaNF_CompraNF = This.IDVendaNF_CompraNF:          End Property
Property Let IDVendaNF_CompraNF(pValue As String):             This.IDVendaNF_CompraNF = pValue:              End Property
Property Get FlagTransf_CompraNF() As String:                   FlagTransf_CompraNF = This.FlagTransf_CompraNF:          End Property
Property Let FlagTransf_CompraNF(pValue As String):             This.FlagTransf_CompraNF = pValue:              End Property


