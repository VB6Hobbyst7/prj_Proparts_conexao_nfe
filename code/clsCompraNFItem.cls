VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompraNFItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private Type ClassType
    ID_CompraNFItem As String
    IDOLD_CompraNFItem As String
    ID_CompraNF_CompraNFItem As String
    ID_CompraNFOLD_CompraNFItem As String
    Item_CompraNFItem As String
    ID_Prod_CompraNFItem As String
    ID_ProdOld_CompraNFItem As String
    ID_Grade_CompraNFItem As String
    Almox_CompraNFItem As String
    QtdFat_CompraNFItem As String
    VUnt_CompraNFItem As String
    TxDesc_CompraNFItem As String
    VUntDesc_CompraNFItem As String
    ICMS_CompraNFItem As String
    ISS_CompraNFItem As String
    IPI_CompraNFItem As String
    ID_NatOp_CompraNFItem As String
    ID_NatOpOLD_CompraNFItem As String
    CFOP_CompraNFItem As String
    ST_CompraNFItem As String
    FlagEst_CompraNFItem As String
    EstDe_CompraNFItem As String
    EstPara_CompraNFItem As String
    DTEmi_CompraNFItem As String
    Esp_CompraNFItem As String
    Série_CompraNFItem As String
    Num_CompraNFItem As String
    Dia_CompraNFItem As String
    UF_CompraNFItem As String
    VTot_CompraNFItem As String
    VCntb_CompraNFItem As String
    BaseCalcICMS_CompraNFItem As String
    VTotBaseCalcICMS_CompraNFItem As String
    DebICMS_CompraNFItem As String
    IseICMS_CompraNFItem As String
    OutICMS_CompraNFItem As String
    BaseCalcIPI_CompraNFItem As String
    DebIPI_CompraNFItem As String
    IseIPI_CompraNFItem As String
    OutIPI_CompraNFItem As String
    Obs_CompraNFItem As String
    TxMLSubsTrib_CompraNFItem As String
    TxIntSubsTrib_CompraNFItem As String
    TxExtSubsTrib_CompraNFItem As String
    BaseCalcICMSSubsTrib_CompraNFItem As String
    VTotICMSSubsTrib_compranfitem As String
    VTotDesc_CompraNFItem As String
    VTotFrete_CompraNFItem As String
    VTotSeg_CompraNFItem As String
    STIPI_CompraNFItem As String
    STPIS_CompraNFItem As String
    STCOFINS_CompraNFItem As String
    nID_CompraNFItem As String
    PIS_CompraNFItem As String
    COFINS_CompraNFItem As String
    VTotBaseCalcPIS_CompraNFItem As String
    VTotBaseCalcCOFINS_CompraNFItem As String
    VTotPIS_CompraNFItem As String
    VTotCOFINS_CompraNFItem As String
    VTotOutDesp_CompraNFItem As String
    VUntCustoSI_CompraNFItem As String
    VTotDebISSRet_CompraNFItem As String
    VTotIseICMS_CompraNFItem As String
    VTotOutICMS_CompraNFItem As String
    SNCredICMS_CompraNFItem As String
    VTotSNCredICMS_CompraNFItem As String
End Type
Private This As ClassType

Property Get ID_CompraNFItem() As String:                   ID_CompraNFItem = This.ID_CompraNFItem:          End Property
Property Let ID_CompraNFItem(pValue As String):             This.ID_CompraNFItem = pValue:              End Property
Property Get IDOLD_CompraNFItem() As String:                   IDOLD_CompraNFItem = This.IDOLD_CompraNFItem:          End Property
Property Let IDOLD_CompraNFItem(pValue As String):             This.IDOLD_CompraNFItem = pValue:              End Property
Property Get ID_CompraNF_CompraNFItem() As String:                   ID_CompraNF_CompraNFItem = This.ID_CompraNF_CompraNFItem:          End Property
Property Let ID_CompraNF_CompraNFItem(pValue As String):             This.ID_CompraNF_CompraNFItem = pValue:              End Property
Property Get ID_CompraNFOLD_CompraNFItem() As String:                   ID_CompraNFOLD_CompraNFItem = This.ID_CompraNFOLD_CompraNFItem:          End Property
Property Let ID_CompraNFOLD_CompraNFItem(pValue As String):             This.ID_CompraNFOLD_CompraNFItem = pValue:              End Property
Property Get Item_CompraNFItem() As String:                   Item_CompraNFItem = This.Item_CompraNFItem:          End Property
Property Let Item_CompraNFItem(pValue As String):             This.Item_CompraNFItem = pValue:              End Property
Property Get ID_Prod_CompraNFItem() As String:                   ID_Prod_CompraNFItem = This.ID_Prod_CompraNFItem:          End Property
Property Let ID_Prod_CompraNFItem(pValue As String):             This.ID_Prod_CompraNFItem = pValue:              End Property
Property Get ID_ProdOld_CompraNFItem() As String:                   ID_ProdOld_CompraNFItem = This.ID_ProdOld_CompraNFItem:          End Property
Property Let ID_ProdOld_CompraNFItem(pValue As String):             This.ID_ProdOld_CompraNFItem = pValue:              End Property
Property Get ID_Grade_CompraNFItem() As String:                   ID_Grade_CompraNFItem = This.ID_Grade_CompraNFItem:          End Property
Property Let ID_Grade_CompraNFItem(pValue As String):             This.ID_Grade_CompraNFItem = pValue:              End Property
Property Get Almox_CompraNFItem() As String:                   Almox_CompraNFItem = This.Almox_CompraNFItem:          End Property
Property Let Almox_CompraNFItem(pValue As String):             This.Almox_CompraNFItem = pValue:              End Property
Property Get QtdFat_CompraNFItem() As String:                   QtdFat_CompraNFItem = This.QtdFat_CompraNFItem:          End Property
Property Let QtdFat_CompraNFItem(pValue As String):             This.QtdFat_CompraNFItem = pValue:              End Property
Property Get VUnt_CompraNFItem() As String:                   VUnt_CompraNFItem = This.VUnt_CompraNFItem:          End Property
Property Let VUnt_CompraNFItem(pValue As String):             This.VUnt_CompraNFItem = pValue:              End Property
Property Get TxDesc_CompraNFItem() As String:                   TxDesc_CompraNFItem = This.TxDesc_CompraNFItem:          End Property
Property Let TxDesc_CompraNFItem(pValue As String):             This.TxDesc_CompraNFItem = pValue:              End Property
Property Get VUntDesc_CompraNFItem() As String:                   VUntDesc_CompraNFItem = This.VUntDesc_CompraNFItem:          End Property
Property Let VUntDesc_CompraNFItem(pValue As String):             This.VUntDesc_CompraNFItem = pValue:              End Property
Property Get ICMS_CompraNFItem() As String:                   ICMS_CompraNFItem = This.ICMS_CompraNFItem:          End Property
Property Let ICMS_CompraNFItem(pValue As String):             This.ICMS_CompraNFItem = pValue:              End Property
Property Get ISS_CompraNFItem() As String:                   ISS_CompraNFItem = This.ISS_CompraNFItem:          End Property
Property Let ISS_CompraNFItem(pValue As String):             This.ISS_CompraNFItem = pValue:              End Property
Property Get IPI_CompraNFItem() As String:                   IPI_CompraNFItem = This.IPI_CompraNFItem:          End Property
Property Let IPI_CompraNFItem(pValue As String):             This.IPI_CompraNFItem = pValue:              End Property
Property Get ID_NatOp_CompraNFItem() As String:                   ID_NatOp_CompraNFItem = This.ID_NatOp_CompraNFItem:          End Property
Property Let ID_NatOp_CompraNFItem(pValue As String):             This.ID_NatOp_CompraNFItem = pValue:              End Property
Property Get ID_NatOpOLD_CompraNFItem() As String:                   ID_NatOpOLD_CompraNFItem = This.ID_NatOpOLD_CompraNFItem:          End Property
Property Let ID_NatOpOLD_CompraNFItem(pValue As String):             This.ID_NatOpOLD_CompraNFItem = pValue:              End Property
Property Get CFOP_CompraNFItem() As String:                   CFOP_CompraNFItem = This.CFOP_CompraNFItem:          End Property
Property Let CFOP_CompraNFItem(pValue As String):             This.CFOP_CompraNFItem = pValue:              End Property
Property Get ST_CompraNFItem() As String:                   ST_CompraNFItem = This.ST_CompraNFItem:          End Property
Property Let ST_CompraNFItem(pValue As String):             This.ST_CompraNFItem = pValue:              End Property
Property Get FlagEst_CompraNFItem() As String:                   FlagEst_CompraNFItem = This.FlagEst_CompraNFItem:          End Property
Property Let FlagEst_CompraNFItem(pValue As String):             This.FlagEst_CompraNFItem = pValue:              End Property
Property Get EstDe_CompraNFItem() As String:                   EstDe_CompraNFItem = This.EstDe_CompraNFItem:          End Property
Property Let EstDe_CompraNFItem(pValue As String):             This.EstDe_CompraNFItem = pValue:              End Property
Property Get EstPara_CompraNFItem() As String:                   EstPara_CompraNFItem = This.EstPara_CompraNFItem:          End Property
Property Let EstPara_CompraNFItem(pValue As String):             This.EstPara_CompraNFItem = pValue:              End Property
Property Get DTEmi_CompraNFItem() As String:                   DTEmi_CompraNFItem = This.DTEmi_CompraNFItem:          End Property
Property Let DTEmi_CompraNFItem(pValue As String):             This.DTEmi_CompraNFItem = pValue:              End Property
Property Get Esp_CompraNFItem() As String:                   Esp_CompraNFItem = This.Esp_CompraNFItem:          End Property
Property Let Esp_CompraNFItem(pValue As String):             This.Esp_CompraNFItem = pValue:              End Property
Property Get Série_CompraNFItem() As String:                   Série_CompraNFItem = This.Série_CompraNFItem:          End Property
Property Let Série_CompraNFItem(pValue As String):             This.Série_CompraNFItem = pValue:              End Property
Property Get Num_CompraNFItem() As String:                   Num_CompraNFItem = This.Num_CompraNFItem:          End Property
Property Let Num_CompraNFItem(pValue As String):             This.Num_CompraNFItem = pValue:              End Property
Property Get Dia_CompraNFItem() As String:                   Dia_CompraNFItem = This.Dia_CompraNFItem:          End Property
Property Let Dia_CompraNFItem(pValue As String):             This.Dia_CompraNFItem = pValue:              End Property
Property Get UF_CompraNFItem() As String:                   UF_CompraNFItem = This.UF_CompraNFItem:          End Property
Property Let UF_CompraNFItem(pValue As String):             This.UF_CompraNFItem = pValue:              End Property
Property Get VTot_CompraNFItem() As String:                   VTot_CompraNFItem = This.VTot_CompraNFItem:          End Property
Property Let VTot_CompraNFItem(pValue As String):             This.VTot_CompraNFItem = pValue:              End Property
Property Get VCntb_CompraNFItem() As String:                   VCntb_CompraNFItem = This.VCntb_CompraNFItem:          End Property
Property Let VCntb_CompraNFItem(pValue As String):             This.VCntb_CompraNFItem = pValue:              End Property
Property Get BaseCalcICMS_CompraNFItem() As String:                   BaseCalcICMS_CompraNFItem = This.BaseCalcICMS_CompraNFItem:          End Property
Property Let BaseCalcICMS_CompraNFItem(pValue As String):             This.BaseCalcICMS_CompraNFItem = pValue:              End Property
Property Get VTotBaseCalcICMS_CompraNFItem() As String:                   VTotBaseCalcICMS_CompraNFItem = This.VTotBaseCalcICMS_CompraNFItem:          End Property
Property Let VTotBaseCalcICMS_CompraNFItem(pValue As String):             This.VTotBaseCalcICMS_CompraNFItem = pValue:              End Property
Property Get DebICMS_CompraNFItem() As String:                   DebICMS_CompraNFItem = This.DebICMS_CompraNFItem:          End Property
Property Let DebICMS_CompraNFItem(pValue As String):             This.DebICMS_CompraNFItem = pValue:              End Property
Property Get IseICMS_CompraNFItem() As String:                   IseICMS_CompraNFItem = This.IseICMS_CompraNFItem:          End Property
Property Let IseICMS_CompraNFItem(pValue As String):             This.IseICMS_CompraNFItem = pValue:              End Property
Property Get OutICMS_CompraNFItem() As String:                   OutICMS_CompraNFItem = This.OutICMS_CompraNFItem:          End Property
Property Let OutICMS_CompraNFItem(pValue As String):             This.OutICMS_CompraNFItem = pValue:              End Property
Property Get BaseCalcIPI_CompraNFItem() As String:                   BaseCalcIPI_CompraNFItem = This.BaseCalcIPI_CompraNFItem:          End Property
Property Let BaseCalcIPI_CompraNFItem(pValue As String):             This.BaseCalcIPI_CompraNFItem = pValue:              End Property
Property Get DebIPI_CompraNFItem() As String:                   DebIPI_CompraNFItem = This.DebIPI_CompraNFItem:          End Property
Property Let DebIPI_CompraNFItem(pValue As String):             This.DebIPI_CompraNFItem = pValue:              End Property
Property Get IseIPI_CompraNFItem() As String:                   IseIPI_CompraNFItem = This.IseIPI_CompraNFItem:          End Property
Property Let IseIPI_CompraNFItem(pValue As String):             This.IseIPI_CompraNFItem = pValue:              End Property
Property Get OutIPI_CompraNFItem() As String:                   OutIPI_CompraNFItem = This.OutIPI_CompraNFItem:          End Property
Property Let OutIPI_CompraNFItem(pValue As String):             This.OutIPI_CompraNFItem = pValue:              End Property
Property Get Obs_CompraNFItem() As String:                   Obs_CompraNFItem = This.Obs_CompraNFItem:          End Property
Property Let Obs_CompraNFItem(pValue As String):             This.Obs_CompraNFItem = pValue:              End Property
Property Get TxMLSubsTrib_CompraNFItem() As String:                   TxMLSubsTrib_CompraNFItem = This.TxMLSubsTrib_CompraNFItem:          End Property
Property Let TxMLSubsTrib_CompraNFItem(pValue As String):             This.TxMLSubsTrib_CompraNFItem = pValue:              End Property
Property Get TxIntSubsTrib_CompraNFItem() As String:                   TxIntSubsTrib_CompraNFItem = This.TxIntSubsTrib_CompraNFItem:          End Property
Property Let TxIntSubsTrib_CompraNFItem(pValue As String):             This.TxIntSubsTrib_CompraNFItem = pValue:              End Property
Property Get TxExtSubsTrib_CompraNFItem() As String:                   TxExtSubsTrib_CompraNFItem = This.TxExtSubsTrib_CompraNFItem:          End Property
Property Let TxExtSubsTrib_CompraNFItem(pValue As String):             This.TxExtSubsTrib_CompraNFItem = pValue:              End Property
Property Get BaseCalcICMSSubsTrib_CompraNFItem() As String:                   BaseCalcICMSSubsTrib_CompraNFItem = This.BaseCalcICMSSubsTrib_CompraNFItem:          End Property
Property Let BaseCalcICMSSubsTrib_CompraNFItem(pValue As String):             This.BaseCalcICMSSubsTrib_CompraNFItem = pValue:              End Property
Property Get VTotICMSSubsTrib_compranfitem() As String:                   VTotICMSSubsTrib_compranfitem = This.VTotICMSSubsTrib_compranfitem:          End Property
Property Let VTotICMSSubsTrib_compranfitem(pValue As String):             This.VTotICMSSubsTrib_compranfitem = pValue:              End Property
Property Get VTotDesc_CompraNFItem() As String:                   VTotDesc_CompraNFItem = This.VTotDesc_CompraNFItem:          End Property
Property Let VTotDesc_CompraNFItem(pValue As String):             This.VTotDesc_CompraNFItem = pValue:              End Property
Property Get VTotFrete_CompraNFItem() As String:                   VTotFrete_CompraNFItem = This.VTotFrete_CompraNFItem:          End Property
Property Let VTotFrete_CompraNFItem(pValue As String):             This.VTotFrete_CompraNFItem = pValue:              End Property
Property Get VTotSeg_CompraNFItem() As String:                   VTotSeg_CompraNFItem = This.VTotSeg_CompraNFItem:          End Property
Property Let VTotSeg_CompraNFItem(pValue As String):             This.VTotSeg_CompraNFItem = pValue:              End Property
Property Get STIPI_CompraNFItem() As String:                   STIPI_CompraNFItem = This.STIPI_CompraNFItem:          End Property
Property Let STIPI_CompraNFItem(pValue As String):             This.STIPI_CompraNFItem = pValue:              End Property
Property Get STPIS_CompraNFItem() As String:                   STPIS_CompraNFItem = This.STPIS_CompraNFItem:          End Property
Property Let STPIS_CompraNFItem(pValue As String):             This.STPIS_CompraNFItem = pValue:              End Property
Property Get STCOFINS_CompraNFItem() As String:                   STCOFINS_CompraNFItem = This.STCOFINS_CompraNFItem:          End Property
Property Let STCOFINS_CompraNFItem(pValue As String):             This.STCOFINS_CompraNFItem = pValue:              End Property
Property Get nID_CompraNFItem() As String:                   nID_CompraNFItem = This.nID_CompraNFItem:          End Property
Property Let nID_CompraNFItem(pValue As String):             This.nID_CompraNFItem = pValue:              End Property
Property Get PIS_CompraNFItem() As String:                   PIS_CompraNFItem = This.PIS_CompraNFItem:          End Property
Property Let PIS_CompraNFItem(pValue As String):             This.PIS_CompraNFItem = pValue:              End Property
Property Get COFINS_CompraNFItem() As String:                   COFINS_CompraNFItem = This.COFINS_CompraNFItem:          End Property
Property Let COFINS_CompraNFItem(pValue As String):             This.COFINS_CompraNFItem = pValue:              End Property
Property Get VTotBaseCalcPIS_CompraNFItem() As String:                   VTotBaseCalcPIS_CompraNFItem = This.VTotBaseCalcPIS_CompraNFItem:          End Property
Property Let VTotBaseCalcPIS_CompraNFItem(pValue As String):             This.VTotBaseCalcPIS_CompraNFItem = pValue:              End Property
Property Get VTotBaseCalcCOFINS_CompraNFItem() As String:                   VTotBaseCalcCOFINS_CompraNFItem = This.VTotBaseCalcCOFINS_CompraNFItem:          End Property
Property Let VTotBaseCalcCOFINS_CompraNFItem(pValue As String):             This.VTotBaseCalcCOFINS_CompraNFItem = pValue:              End Property
Property Get VTotPIS_CompraNFItem() As String:                   VTotPIS_CompraNFItem = This.VTotPIS_CompraNFItem:          End Property
Property Let VTotPIS_CompraNFItem(pValue As String):             This.VTotPIS_CompraNFItem = pValue:              End Property
Property Get VTotCOFINS_CompraNFItem() As String:                   VTotCOFINS_CompraNFItem = This.VTotCOFINS_CompraNFItem:          End Property
Property Let VTotCOFINS_CompraNFItem(pValue As String):             This.VTotCOFINS_CompraNFItem = pValue:              End Property
Property Get VTotOutDesp_CompraNFItem() As String:                   VTotOutDesp_CompraNFItem = This.VTotOutDesp_CompraNFItem:          End Property
Property Let VTotOutDesp_CompraNFItem(pValue As String):             This.VTotOutDesp_CompraNFItem = pValue:              End Property
Property Get VUntCustoSI_CompraNFItem() As String:                   VUntCustoSI_CompraNFItem = This.VUntCustoSI_CompraNFItem:          End Property
Property Let VUntCustoSI_CompraNFItem(pValue As String):             This.VUntCustoSI_CompraNFItem = pValue:              End Property
Property Get VTotDebISSRet_CompraNFItem() As String:                   VTotDebISSRet_CompraNFItem = This.VTotDebISSRet_CompraNFItem:          End Property
Property Let VTotDebISSRet_CompraNFItem(pValue As String):             This.VTotDebISSRet_CompraNFItem = pValue:              End Property
Property Get VTotIseICMS_CompraNFItem() As String:                   VTotIseICMS_CompraNFItem = This.VTotIseICMS_CompraNFItem:          End Property
Property Let VTotIseICMS_CompraNFItem(pValue As String):             This.VTotIseICMS_CompraNFItem = pValue:              End Property
Property Get VTotOutICMS_CompraNFItem() As String:                   VTotOutICMS_CompraNFItem = This.VTotOutICMS_CompraNFItem:          End Property
Property Let VTotOutICMS_CompraNFItem(pValue As String):             This.VTotOutICMS_CompraNFItem = pValue:              End Property
Property Get SNCredICMS_CompraNFItem() As String:                   SNCredICMS_CompraNFItem = This.SNCredICMS_CompraNFItem:          End Property
Property Let SNCredICMS_CompraNFItem(pValue As String):             This.SNCredICMS_CompraNFItem = pValue:              End Property
Property Get VTotSNCredICMS_CompraNFItem() As String:                   VTotSNCredICMS_CompraNFItem = This.VTotSNCredICMS_CompraNFItem:          End Property
Property Let VTotSNCredICMS_CompraNFItem(pValue As String):             This.VTotSNCredICMS_CompraNFItem = pValue:              End Property


