VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompraNFItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type ClassType
	IDOLD_CompraNFItem As Long
	ID_CompraNF_CompraNFItem As Long
	ID_CompraNFOLD_CompraNFItem As Long
	Item_CompraNFItem As Long
	ID_Prod_CompraNFItem As Long
	ID_ProdOld_CompraNFItem As Long
	ID_Grade_CompraNFItem As Long
	Almox_CompraNFItem As Long
	QtdFat_CompraNFItem As Long
	VUnt_CompraNFItem As Double
	TxDesc_CompraNFItem As Double
	VUntDesc_CompraNFItem As Double
	ICMS_CompraNFItem As Double
	ISS_CompraNFItem As Double
	IPI_CompraNFItem As Double
	ID_NatOp_CompraNFItem As Long
	ID_NatOpOLD_CompraNFItem As Long
	CFOP_CompraNFItem As String
	ST_CompraNFItem As String
	FlagEst_CompraNFItem As Byte
	EstDe_CompraNFItem As String
	EstPara_CompraNFItem As String
	DTEmi_CompraNFItem As Date
	Esp_CompraNFItem As String
	Série_CompraNFItem As String
	Num_CompraNFItem As String
	Dia_CompraNFItem As String
	UF_CompraNFItem As String
	VTot_CompraNFItem As Double
	VCntb_CompraNFItem As Double
	BaseCalcICMS_CompraNFItem As Double
	VTotBaseCalcICMS_CompraNFItem As Double
	DebICMS_CompraNFItem As Double
	IseICMS_CompraNFItem As Double
	OutICMS_CompraNFItem As Double
	BaseCalcIPI_CompraNFItem As Double
	DebIPI_CompraNFItem As Double
	IseIPI_CompraNFItem As Double
	OutIPI_CompraNFItem As Double
	Obs_CompraNFItem As String
	TxMLSubsTrib_CompraNFItem As Double
	TxIntSubsTrib_CompraNFItem As Double
	TxExtSubsTrib_CompraNFItem As Double
	BaseCalcICMSSubsTrib_CompraNFItem As Double
	VTotICMSSubsTrib_compranfitem As Double
	VTotDesc_CompraNFItem As Double
	VTotFrete_CompraNFItem As Double
	VTotSeg_CompraNFItem As Double
	STIPI_CompraNFItem As String
	STPIS_CompraNFItem As String
	STCOFINS_CompraNFItem As String
	nID_CompraNFItem As String
	PIS_CompraNFItem As Double
	COFINS_CompraNFItem As Double
	VTotBaseCalcPIS_CompraNFItem As Double
	VTotBaseCalcCOFINS_CompraNFItem As Double
	VTotPIS_CompraNFItem As Double
	VTotCOFINS_CompraNFItem As Double
	VTotOutDesp_CompraNFItem As Double
	VUntCustoSI_CompraNFItem As Double
	VTotDebISSRet_CompraNFItem As Double
	VTotIseICMS_CompraNFItem As Double
	VTotOutICMS_CompraNFItem As Double
	SNCredICMS_CompraNFItem As Double
	VTotSNCredICMS_CompraNFItem As Double
End Type
Private This As ClassType


Property Get IDOLD_CompraNFItem() As Long:IDOLD_CompraNFItem = This.IDOLD_CompraNFItem:End Property
Property Let IDOLD_CompraNFItem(pValue As Long):This.IDOLD_CompraNFItem = pValue:End Property
Property Get ID_CompraNF_CompraNFItem() As Long:ID_CompraNF_CompraNFItem = This.ID_CompraNF_CompraNFItem:End Property
Property Let ID_CompraNF_CompraNFItem(pValue As Long):This.ID_CompraNF_CompraNFItem = pValue:End Property
Property Get ID_CompraNFOLD_CompraNFItem() As Long:ID_CompraNFOLD_CompraNFItem = This.ID_CompraNFOLD_CompraNFItem:End Property
Property Let ID_CompraNFOLD_CompraNFItem(pValue As Long):This.ID_CompraNFOLD_CompraNFItem = pValue:End Property
Property Get Item_CompraNFItem() As Long:Item_CompraNFItem = This.Item_CompraNFItem:End Property
Property Let Item_CompraNFItem(pValue As Long):This.Item_CompraNFItem = pValue:End Property
Property Get ID_Prod_CompraNFItem() As Long:ID_Prod_CompraNFItem = This.ID_Prod_CompraNFItem:End Property
Property Let ID_Prod_CompraNFItem(pValue As Long):This.ID_Prod_CompraNFItem = pValue:End Property
Property Get ID_ProdOld_CompraNFItem() As Long:ID_ProdOld_CompraNFItem = This.ID_ProdOld_CompraNFItem:End Property
Property Let ID_ProdOld_CompraNFItem(pValue As Long):This.ID_ProdOld_CompraNFItem = pValue:End Property
Property Get ID_Grade_CompraNFItem() As Long:ID_Grade_CompraNFItem = This.ID_Grade_CompraNFItem:End Property
Property Let ID_Grade_CompraNFItem(pValue As Long):This.ID_Grade_CompraNFItem = pValue:End Property
Property Get Almox_CompraNFItem() As Long:Almox_CompraNFItem = This.Almox_CompraNFItem:End Property
Property Let Almox_CompraNFItem(pValue As Long):This.Almox_CompraNFItem = pValue:End Property
Property Get QtdFat_CompraNFItem() As Long:QtdFat_CompraNFItem = This.QtdFat_CompraNFItem:End Property
Property Let QtdFat_CompraNFItem(pValue As Long):This.QtdFat_CompraNFItem = pValue:End Property
Property Get VUnt_CompraNFItem() As Double:VUnt_CompraNFItem = This.VUnt_CompraNFItem:End Property
Property Let VUnt_CompraNFItem(pValue As Double):This.VUnt_CompraNFItem = pValue:End Property
Property Get TxDesc_CompraNFItem() As Double:TxDesc_CompraNFItem = This.TxDesc_CompraNFItem:End Property
Property Let TxDesc_CompraNFItem(pValue As Double):This.TxDesc_CompraNFItem = pValue:End Property
Property Get VUntDesc_CompraNFItem() As Double:VUntDesc_CompraNFItem = This.VUntDesc_CompraNFItem:End Property
Property Let VUntDesc_CompraNFItem(pValue As Double):This.VUntDesc_CompraNFItem = pValue:End Property
Property Get ICMS_CompraNFItem() As Double:ICMS_CompraNFItem = This.ICMS_CompraNFItem:End Property
Property Let ICMS_CompraNFItem(pValue As Double):This.ICMS_CompraNFItem = pValue:End Property
Property Get ISS_CompraNFItem() As Double:ISS_CompraNFItem = This.ISS_CompraNFItem:End Property
Property Let ISS_CompraNFItem(pValue As Double):This.ISS_CompraNFItem = pValue:End Property
Property Get IPI_CompraNFItem() As Double:IPI_CompraNFItem = This.IPI_CompraNFItem:End Property
Property Let IPI_CompraNFItem(pValue As Double):This.IPI_CompraNFItem = pValue:End Property
Property Get ID_NatOp_CompraNFItem() As Long:ID_NatOp_CompraNFItem = This.ID_NatOp_CompraNFItem:End Property
Property Let ID_NatOp_CompraNFItem(pValue As Long):This.ID_NatOp_CompraNFItem = pValue:End Property
Property Get ID_NatOpOLD_CompraNFItem() As Long:ID_NatOpOLD_CompraNFItem = This.ID_NatOpOLD_CompraNFItem:End Property
Property Let ID_NatOpOLD_CompraNFItem(pValue As Long):This.ID_NatOpOLD_CompraNFItem = pValue:End Property
Property Get CFOP_CompraNFItem() As String:CFOP_CompraNFItem = This.CFOP_CompraNFItem:End Property
Property Let CFOP_CompraNFItem(pValue As String):This.CFOP_CompraNFItem = pValue:End Property
Property Get ST_CompraNFItem() As String:ST_CompraNFItem = This.ST_CompraNFItem:End Property
Property Let ST_CompraNFItem(pValue As String):This.ST_CompraNFItem = pValue:End Property
Property Get FlagEst_CompraNFItem() As Byte:FlagEst_CompraNFItem = This.FlagEst_CompraNFItem:End Property
Property Let FlagEst_CompraNFItem(pValue As Byte):This.FlagEst_CompraNFItem = pValue:End Property
Property Get EstDe_CompraNFItem() As String:EstDe_CompraNFItem = This.EstDe_CompraNFItem:End Property
Property Let EstDe_CompraNFItem(pValue As String):This.EstDe_CompraNFItem = pValue:End Property
Property Get EstPara_CompraNFItem() As String:EstPara_CompraNFItem = This.EstPara_CompraNFItem:End Property
Property Let EstPara_CompraNFItem(pValue As String):This.EstPara_CompraNFItem = pValue:End Property
Property Get DTEmi_CompraNFItem() As Date:DTEmi_CompraNFItem = This.DTEmi_CompraNFItem:End Property
Property Let DTEmi_CompraNFItem(pValue As Date):This.DTEmi_CompraNFItem = pValue:End Property
Property Get Esp_CompraNFItem() As String:Esp_CompraNFItem = This.Esp_CompraNFItem:End Property
Property Let Esp_CompraNFItem(pValue As String):This.Esp_CompraNFItem = pValue:End Property
Property Get Série_CompraNFItem() As String:Série_CompraNFItem = This.Série_CompraNFItem:End Property
Property Let Série_CompraNFItem(pValue As String):This.Série_CompraNFItem = pValue:End Property
Property Get Num_CompraNFItem() As String:Num_CompraNFItem = This.Num_CompraNFItem:End Property
Property Let Num_CompraNFItem(pValue As String):This.Num_CompraNFItem = pValue:End Property
Property Get Dia_CompraNFItem() As String:Dia_CompraNFItem = This.Dia_CompraNFItem:End Property
Property Let Dia_CompraNFItem(pValue As String):This.Dia_CompraNFItem = pValue:End Property
Property Get UF_CompraNFItem() As String:UF_CompraNFItem = This.UF_CompraNFItem:End Property
Property Let UF_CompraNFItem(pValue As String):This.UF_CompraNFItem = pValue:End Property
Property Get VTot_CompraNFItem() As Double:VTot_CompraNFItem = This.VTot_CompraNFItem:End Property
Property Let VTot_CompraNFItem(pValue As Double):This.VTot_CompraNFItem = pValue:End Property
Property Get VCntb_CompraNFItem() As Double:VCntb_CompraNFItem = This.VCntb_CompraNFItem:End Property
Property Let VCntb_CompraNFItem(pValue As Double):This.VCntb_CompraNFItem = pValue:End Property
Property Get BaseCalcICMS_CompraNFItem() As Double:BaseCalcICMS_CompraNFItem = This.BaseCalcICMS_CompraNFItem:End Property
Property Let BaseCalcICMS_CompraNFItem(pValue As Double):This.BaseCalcICMS_CompraNFItem = pValue:End Property
Property Get VTotBaseCalcICMS_CompraNFItem() As Double:VTotBaseCalcICMS_CompraNFItem = This.VTotBaseCalcICMS_CompraNFItem:End Property
Property Let VTotBaseCalcICMS_CompraNFItem(pValue As Double):This.VTotBaseCalcICMS_CompraNFItem = pValue:End Property
Property Get DebICMS_CompraNFItem() As Double:DebICMS_CompraNFItem = This.DebICMS_CompraNFItem:End Property
Property Let DebICMS_CompraNFItem(pValue As Double):This.DebICMS_CompraNFItem = pValue:End Property
Property Get IseICMS_CompraNFItem() As Double:IseICMS_CompraNFItem = This.IseICMS_CompraNFItem:End Property
Property Let IseICMS_CompraNFItem(pValue As Double):This.IseICMS_CompraNFItem = pValue:End Property
Property Get OutICMS_CompraNFItem() As Double:OutICMS_CompraNFItem = This.OutICMS_CompraNFItem:End Property
Property Let OutICMS_CompraNFItem(pValue As Double):This.OutICMS_CompraNFItem = pValue:End Property
Property Get BaseCalcIPI_CompraNFItem() As Double:BaseCalcIPI_CompraNFItem = This.BaseCalcIPI_CompraNFItem:End Property
Property Let BaseCalcIPI_CompraNFItem(pValue As Double):This.BaseCalcIPI_CompraNFItem = pValue:End Property
Property Get DebIPI_CompraNFItem() As Double:DebIPI_CompraNFItem = This.DebIPI_CompraNFItem:End Property
Property Let DebIPI_CompraNFItem(pValue As Double):This.DebIPI_CompraNFItem = pValue:End Property
Property Get IseIPI_CompraNFItem() As Double:IseIPI_CompraNFItem = This.IseIPI_CompraNFItem:End Property
Property Let IseIPI_CompraNFItem(pValue As Double):This.IseIPI_CompraNFItem = pValue:End Property
Property Get OutIPI_CompraNFItem() As Double:OutIPI_CompraNFItem = This.OutIPI_CompraNFItem:End Property
Property Let OutIPI_CompraNFItem(pValue As Double):This.OutIPI_CompraNFItem = pValue:End Property
Property Get Obs_CompraNFItem() As String:Obs_CompraNFItem = This.Obs_CompraNFItem:End Property
Property Let Obs_CompraNFItem(pValue As String):This.Obs_CompraNFItem = pValue:End Property
Property Get TxMLSubsTrib_CompraNFItem() As Double:TxMLSubsTrib_CompraNFItem = This.TxMLSubsTrib_CompraNFItem:End Property
Property Let TxMLSubsTrib_CompraNFItem(pValue As Double):This.TxMLSubsTrib_CompraNFItem = pValue:End Property
Property Get TxIntSubsTrib_CompraNFItem() As Double:TxIntSubsTrib_CompraNFItem = This.TxIntSubsTrib_CompraNFItem:End Property
Property Let TxIntSubsTrib_CompraNFItem(pValue As Double):This.TxIntSubsTrib_CompraNFItem = pValue:End Property
Property Get TxExtSubsTrib_CompraNFItem() As Double:TxExtSubsTrib_CompraNFItem = This.TxExtSubsTrib_CompraNFItem:End Property
Property Let TxExtSubsTrib_CompraNFItem(pValue As Double):This.TxExtSubsTrib_CompraNFItem = pValue:End Property
Property Get BaseCalcICMSSubsTrib_CompraNFItem() As Double:BaseCalcICMSSubsTrib_CompraNFItem = This.BaseCalcICMSSubsTrib_CompraNFItem:End Property
Property Let BaseCalcICMSSubsTrib_CompraNFItem(pValue As Double):This.BaseCalcICMSSubsTrib_CompraNFItem = pValue:End Property
Property Get VTotICMSSubsTrib_compranfitem() As Double:VTotICMSSubsTrib_compranfitem = This.VTotICMSSubsTrib_compranfitem:End Property
Property Let VTotICMSSubsTrib_compranfitem(pValue As Double):This.VTotICMSSubsTrib_compranfitem = pValue:End Property
Property Get VTotDesc_CompraNFItem() As Double:VTotDesc_CompraNFItem = This.VTotDesc_CompraNFItem:End Property
Property Let VTotDesc_CompraNFItem(pValue As Double):This.VTotDesc_CompraNFItem = pValue:End Property
Property Get VTotFrete_CompraNFItem() As Double:VTotFrete_CompraNFItem = This.VTotFrete_CompraNFItem:End Property
Property Let VTotFrete_CompraNFItem(pValue As Double):This.VTotFrete_CompraNFItem = pValue:End Property
Property Get VTotSeg_CompraNFItem() As Double:VTotSeg_CompraNFItem = This.VTotSeg_CompraNFItem:End Property
Property Let VTotSeg_CompraNFItem(pValue As Double):This.VTotSeg_CompraNFItem = pValue:End Property
Property Get STIPI_CompraNFItem() As String:STIPI_CompraNFItem = This.STIPI_CompraNFItem:End Property
Property Let STIPI_CompraNFItem(pValue As String):This.STIPI_CompraNFItem = pValue:End Property
Property Get STPIS_CompraNFItem() As String:STPIS_CompraNFItem = This.STPIS_CompraNFItem:End Property
Property Let STPIS_CompraNFItem(pValue As String):This.STPIS_CompraNFItem = pValue:End Property
Property Get STCOFINS_CompraNFItem() As String:STCOFINS_CompraNFItem = This.STCOFINS_CompraNFItem:End Property
Property Let STCOFINS_CompraNFItem(pValue As String):This.STCOFINS_CompraNFItem = pValue:End Property
Property Get nID_CompraNFItem() As String:nID_CompraNFItem = This.nID_CompraNFItem:End Property
Property Let nID_CompraNFItem(pValue As String):This.nID_CompraNFItem = pValue:End Property
Property Get PIS_CompraNFItem() As Double:PIS_CompraNFItem = This.PIS_CompraNFItem:End Property
Property Let PIS_CompraNFItem(pValue As Double):This.PIS_CompraNFItem = pValue:End Property
Property Get COFINS_CompraNFItem() As Double:COFINS_CompraNFItem = This.COFINS_CompraNFItem:End Property
Property Let COFINS_CompraNFItem(pValue As Double):This.COFINS_CompraNFItem = pValue:End Property
Property Get VTotBaseCalcPIS_CompraNFItem() As Double:VTotBaseCalcPIS_CompraNFItem = This.VTotBaseCalcPIS_CompraNFItem:End Property
Property Let VTotBaseCalcPIS_CompraNFItem(pValue As Double):This.VTotBaseCalcPIS_CompraNFItem = pValue:End Property
Property Get VTotBaseCalcCOFINS_CompraNFItem() As Double:VTotBaseCalcCOFINS_CompraNFItem = This.VTotBaseCalcCOFINS_CompraNFItem:End Property
Property Let VTotBaseCalcCOFINS_CompraNFItem(pValue As Double):This.VTotBaseCalcCOFINS_CompraNFItem = pValue:End Property
Property Get VTotPIS_CompraNFItem() As Double:VTotPIS_CompraNFItem = This.VTotPIS_CompraNFItem:End Property
Property Let VTotPIS_CompraNFItem(pValue As Double):This.VTotPIS_CompraNFItem = pValue:End Property
Property Get VTotCOFINS_CompraNFItem() As Double:VTotCOFINS_CompraNFItem = This.VTotCOFINS_CompraNFItem:End Property
Property Let VTotCOFINS_CompraNFItem(pValue As Double):This.VTotCOFINS_CompraNFItem = pValue:End Property
Property Get VTotOutDesp_CompraNFItem() As Double:VTotOutDesp_CompraNFItem = This.VTotOutDesp_CompraNFItem:End Property
Property Let VTotOutDesp_CompraNFItem(pValue As Double):This.VTotOutDesp_CompraNFItem = pValue:End Property
Property Get VUntCustoSI_CompraNFItem() As Double:VUntCustoSI_CompraNFItem = This.VUntCustoSI_CompraNFItem:End Property
Property Let VUntCustoSI_CompraNFItem(pValue As Double):This.VUntCustoSI_CompraNFItem = pValue:End Property
Property Get VTotDebISSRet_CompraNFItem() As Double:VTotDebISSRet_CompraNFItem = This.VTotDebISSRet_CompraNFItem:End Property
Property Let VTotDebISSRet_CompraNFItem(pValue As Double):This.VTotDebISSRet_CompraNFItem = pValue:End Property
Property Get VTotIseICMS_CompraNFItem() As Double:VTotIseICMS_CompraNFItem = This.VTotIseICMS_CompraNFItem:End Property
Property Let VTotIseICMS_CompraNFItem(pValue As Double):This.VTotIseICMS_CompraNFItem = pValue:End Property
Property Get VTotOutICMS_CompraNFItem() As Double:VTotOutICMS_CompraNFItem = This.VTotOutICMS_CompraNFItem:End Property
Property Let VTotOutICMS_CompraNFItem(pValue As Double):This.VTotOutICMS_CompraNFItem = pValue:End Property
Property Get SNCredICMS_CompraNFItem() As Double:SNCredICMS_CompraNFItem = This.SNCredICMS_CompraNFItem:End Property
Property Let SNCredICMS_CompraNFItem(pValue As Double):This.SNCredICMS_CompraNFItem = pValue:End Property
Property Get VTotSNCredICMS_CompraNFItem() As Double:VTotSNCredICMS_CompraNFItem = This.VTotSNCredICMS_CompraNFItem:End Property
Property Let VTotSNCredICMS_CompraNFItem(pValue As Double):This.VTotSNCredICMS_CompraNFItem = pValue:End Property


sub cadastrar()
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblCompraNFItem")


With qdf
	.Parameters("IDOLD_CompraNFItem") = IDOLD_CompraNFItem
	.Parameters("ID_CompraNF_CompraNFItem") = ID_CompraNF_CompraNFItem
	.Parameters("ID_CompraNFOLD_CompraNFItem") = ID_CompraNFOLD_CompraNFItem
	.Parameters("Item_CompraNFItem") = Item_CompraNFItem
	.Parameters("ID_Prod_CompraNFItem") = ID_Prod_CompraNFItem
	.Parameters("ID_ProdOld_CompraNFItem") = ID_ProdOld_CompraNFItem
	.Parameters("ID_Grade_CompraNFItem") = ID_Grade_CompraNFItem
	.Parameters("Almox_CompraNFItem") = Almox_CompraNFItem
	.Parameters("QtdFat_CompraNFItem") = QtdFat_CompraNFItem
	.Parameters("VUnt_CompraNFItem") = VUnt_CompraNFItem
	.Parameters("TxDesc_CompraNFItem") = TxDesc_CompraNFItem
	.Parameters("VUntDesc_CompraNFItem") = VUntDesc_CompraNFItem
	.Parameters("ICMS_CompraNFItem") = ICMS_CompraNFItem
	.Parameters("ISS_CompraNFItem") = ISS_CompraNFItem
	.Parameters("IPI_CompraNFItem") = IPI_CompraNFItem
	.Parameters("ID_NatOp_CompraNFItem") = ID_NatOp_CompraNFItem
	.Parameters("ID_NatOpOLD_CompraNFItem") = ID_NatOpOLD_CompraNFItem
	.Parameters("CFOP_CompraNFItem") = CFOP_CompraNFItem
	.Parameters("ST_CompraNFItem") = ST_CompraNFItem
	.Parameters("FlagEst_CompraNFItem") = FlagEst_CompraNFItem
	.Parameters("EstDe_CompraNFItem") = EstDe_CompraNFItem
	.Parameters("EstPara_CompraNFItem") = EstPara_CompraNFItem
	.Parameters("DTEmi_CompraNFItem") = DTEmi_CompraNFItem
	.Parameters("Esp_CompraNFItem") = Esp_CompraNFItem
	.Parameters("Série_CompraNFItem") = Série_CompraNFItem
	.Parameters("Num_CompraNFItem") = Num_CompraNFItem
	.Parameters("Dia_CompraNFItem") = Dia_CompraNFItem
	.Parameters("UF_CompraNFItem") = UF_CompraNFItem
	.Parameters("VTot_CompraNFItem") = VTot_CompraNFItem
	.Parameters("VCntb_CompraNFItem") = VCntb_CompraNFItem
	.Parameters("BaseCalcICMS_CompraNFItem") = BaseCalcICMS_CompraNFItem
	.Parameters("VTotBaseCalcICMS_CompraNFItem") = VTotBaseCalcICMS_CompraNFItem
	.Parameters("DebICMS_CompraNFItem") = DebICMS_CompraNFItem
	.Parameters("IseICMS_CompraNFItem") = IseICMS_CompraNFItem
	.Parameters("OutICMS_CompraNFItem") = OutICMS_CompraNFItem
	.Parameters("BaseCalcIPI_CompraNFItem") = BaseCalcIPI_CompraNFItem
	.Parameters("DebIPI_CompraNFItem") = DebIPI_CompraNFItem
	.Parameters("IseIPI_CompraNFItem") = IseIPI_CompraNFItem
	.Parameters("OutIPI_CompraNFItem") = OutIPI_CompraNFItem
	.Parameters("Obs_CompraNFItem") = Obs_CompraNFItem
	.Parameters("TxMLSubsTrib_CompraNFItem") = TxMLSubsTrib_CompraNFItem
	.Parameters("TxIntSubsTrib_CompraNFItem") = TxIntSubsTrib_CompraNFItem
	.Parameters("TxExtSubsTrib_CompraNFItem") = TxExtSubsTrib_CompraNFItem
	.Parameters("BaseCalcICMSSubsTrib_CompraNFItem") = BaseCalcICMSSubsTrib_CompraNFItem
	.Parameters("VTotICMSSubsTrib_compranfitem") = VTotICMSSubsTrib_compranfitem
	.Parameters("VTotDesc_CompraNFItem") = VTotDesc_CompraNFItem
	.Parameters("VTotFrete_CompraNFItem") = VTotFrete_CompraNFItem
	.Parameters("VTotSeg_CompraNFItem") = VTotSeg_CompraNFItem
	.Parameters("STIPI_CompraNFItem") = STIPI_CompraNFItem
	.Parameters("STPIS_CompraNFItem") = STPIS_CompraNFItem
	.Parameters("STCOFINS_CompraNFItem") = STCOFINS_CompraNFItem
	.Parameters("nID_CompraNFItem") = nID_CompraNFItem
	.Parameters("PIS_CompraNFItem") = PIS_CompraNFItem
	.Parameters("COFINS_CompraNFItem") = COFINS_CompraNFItem
	.Parameters("VTotBaseCalcPIS_CompraNFItem") = VTotBaseCalcPIS_CompraNFItem
	.Parameters("VTotBaseCalcCOFINS_CompraNFItem") = VTotBaseCalcCOFINS_CompraNFItem
	.Parameters("VTotPIS_CompraNFItem") = VTotPIS_CompraNFItem
	.Parameters("VTotCOFINS_CompraNFItem") = VTotCOFINS_CompraNFItem
	.Parameters("VTotOutDesp_CompraNFItem") = VTotOutDesp_CompraNFItem
	.Parameters("VUntCustoSI_CompraNFItem") = VUntCustoSI_CompraNFItem
	.Parameters("VTotDebISSRet_CompraNFItem") = VTotDebISSRet_CompraNFItem
	.Parameters("VTotIseICMS_CompraNFItem") = VTotIseICMS_CompraNFItem
	.Parameters("VTotOutICMS_CompraNFItem") = VTotOutICMS_CompraNFItem
	.Parameters("SNCredICMS_CompraNFItem") = SNCredICMS_CompraNFItem
	.Parameters("VTotSNCredICMS_CompraNFItem") = VTotSNCredICMS_CompraNFItem
	.Execute
End With
End Sub

