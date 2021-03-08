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


Private This As ClassType


Private Type ClassType
	ID_CompraNFItem As Long
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


Property Get ID_CompraNFItem() As Long:ID_CompraNFItem = This.ID_CompraNFItem:End Property
Property Let ID_CompraNFItem(pValue As Long):This.ID_CompraNFItem = pValue:End Property
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


sub cadastrar(obj as clsCompraNFItem)
Dim db As DAO.Database: Set db = CurrentDb
Dim qdf As DAO.QueryDef: Set qdf = db.QueryDefs("INSERT_tblCompraNFItem")


With qdf
	.Parameters("strID_CompraNFItem") = obj.ID_CompraNFItem
	.Parameters("strIDOLD_CompraNFItem") = obj.IDOLD_CompraNFItem
	.Parameters("strID_CompraNF_CompraNFItem") = obj.ID_CompraNF_CompraNFItem
	.Parameters("strID_CompraNFOLD_CompraNFItem") = obj.ID_CompraNFOLD_CompraNFItem
	.Parameters("strItem_CompraNFItem") = obj.Item_CompraNFItem
	.Parameters("strID_Prod_CompraNFItem") = obj.ID_Prod_CompraNFItem
	.Parameters("strID_ProdOld_CompraNFItem") = obj.ID_ProdOld_CompraNFItem
	.Parameters("strID_Grade_CompraNFItem") = obj.ID_Grade_CompraNFItem
	.Parameters("strAlmox_CompraNFItem") = obj.Almox_CompraNFItem
	.Parameters("strQtdFat_CompraNFItem") = obj.QtdFat_CompraNFItem
	.Parameters("strVUnt_CompraNFItem") = obj.VUnt_CompraNFItem
	.Parameters("strTxDesc_CompraNFItem") = obj.TxDesc_CompraNFItem
	.Parameters("strVUntDesc_CompraNFItem") = obj.VUntDesc_CompraNFItem
	.Parameters("strICMS_CompraNFItem") = obj.ICMS_CompraNFItem
	.Parameters("strISS_CompraNFItem") = obj.ISS_CompraNFItem
	.Parameters("strIPI_CompraNFItem") = obj.IPI_CompraNFItem
	.Parameters("strID_NatOp_CompraNFItem") = obj.ID_NatOp_CompraNFItem
	.Parameters("strID_NatOpOLD_CompraNFItem") = obj.ID_NatOpOLD_CompraNFItem
	.Parameters("strCFOP_CompraNFItem") = obj.CFOP_CompraNFItem
	.Parameters("strST_CompraNFItem") = obj.ST_CompraNFItem
	.Parameters("strFlagEst_CompraNFItem") = obj.FlagEst_CompraNFItem
	.Parameters("strEstDe_CompraNFItem") = obj.EstDe_CompraNFItem
	.Parameters("strEstPara_CompraNFItem") = obj.EstPara_CompraNFItem
	.Parameters("strDTEmi_CompraNFItem") = obj.DTEmi_CompraNFItem
	.Parameters("strEsp_CompraNFItem") = obj.Esp_CompraNFItem
	.Parameters("strSérie_CompraNFItem") = obj.Série_CompraNFItem
	.Parameters("strNum_CompraNFItem") = obj.Num_CompraNFItem
	.Parameters("strDia_CompraNFItem") = obj.Dia_CompraNFItem
	.Parameters("strUF_CompraNFItem") = obj.UF_CompraNFItem
	.Parameters("strVTot_CompraNFItem") = obj.VTot_CompraNFItem
	.Parameters("strVCntb_CompraNFItem") = obj.VCntb_CompraNFItem
	.Parameters("strBaseCalcICMS_CompraNFItem") = obj.BaseCalcICMS_CompraNFItem
	.Parameters("strVTotBaseCalcICMS_CompraNFItem") = obj.VTotBaseCalcICMS_CompraNFItem
	.Parameters("strDebICMS_CompraNFItem") = obj.DebICMS_CompraNFItem
	.Parameters("strIseICMS_CompraNFItem") = obj.IseICMS_CompraNFItem
	.Parameters("strOutICMS_CompraNFItem") = obj.OutICMS_CompraNFItem
	.Parameters("strBaseCalcIPI_CompraNFItem") = obj.BaseCalcIPI_CompraNFItem
	.Parameters("strDebIPI_CompraNFItem") = obj.DebIPI_CompraNFItem
	.Parameters("strIseIPI_CompraNFItem") = obj.IseIPI_CompraNFItem
	.Parameters("strOutIPI_CompraNFItem") = obj.OutIPI_CompraNFItem
	.Parameters("strObs_CompraNFItem") = obj.Obs_CompraNFItem
	.Parameters("strTxMLSubsTrib_CompraNFItem") = obj.TxMLSubsTrib_CompraNFItem
	.Parameters("strTxIntSubsTrib_CompraNFItem") = obj.TxIntSubsTrib_CompraNFItem
	.Parameters("strTxExtSubsTrib_CompraNFItem") = obj.TxExtSubsTrib_CompraNFItem
	.Parameters("strBaseCalcICMSSubsTrib_CompraNFItem") = obj.BaseCalcICMSSubsTrib_CompraNFItem
	.Parameters("strVTotICMSSubsTrib_compranfitem") = obj.VTotICMSSubsTrib_compranfitem
	.Parameters("strVTotDesc_CompraNFItem") = obj.VTotDesc_CompraNFItem
	.Parameters("strVTotFrete_CompraNFItem") = obj.VTotFrete_CompraNFItem
	.Parameters("strVTotSeg_CompraNFItem") = obj.VTotSeg_CompraNFItem
	.Parameters("strSTIPI_CompraNFItem") = obj.STIPI_CompraNFItem
	.Parameters("strSTPIS_CompraNFItem") = obj.STPIS_CompraNFItem
	.Parameters("strSTCOFINS_CompraNFItem") = obj.STCOFINS_CompraNFItem
	.Parameters("strnID_CompraNFItem") = obj.nID_CompraNFItem
	.Parameters("strPIS_CompraNFItem") = obj.PIS_CompraNFItem
	.Parameters("strCOFINS_CompraNFItem") = obj.COFINS_CompraNFItem
	.Parameters("strVTotBaseCalcPIS_CompraNFItem") = obj.VTotBaseCalcPIS_CompraNFItem
	.Parameters("strVTotBaseCalcCOFINS_CompraNFItem") = obj.VTotBaseCalcCOFINS_CompraNFItem
	.Parameters("strVTotPIS_CompraNFItem") = obj.VTotPIS_CompraNFItem
	.Parameters("strVTotCOFINS_CompraNFItem") = obj.VTotCOFINS_CompraNFItem
	.Parameters("strVTotOutDesp_CompraNFItem") = obj.VTotOutDesp_CompraNFItem
	.Parameters("strVUntCustoSI_CompraNFItem") = obj.VUntCustoSI_CompraNFItem
	.Parameters("strVTotDebISSRet_CompraNFItem") = obj.VTotDebISSRet_CompraNFItem
	.Parameters("strVTotIseICMS_CompraNFItem") = obj.VTotIseICMS_CompraNFItem
	.Parameters("strVTotOutICMS_CompraNFItem") = obj.VTotOutICMS_CompraNFItem
	.Parameters("strSNCredICMS_CompraNFItem") = obj.SNCredICMS_CompraNFItem
	.Parameters("strVTotSNCredICMS_CompraNFItem") = obj.VTotSNCredICMS_CompraNFItem
	.Execute
End With
End Sub

