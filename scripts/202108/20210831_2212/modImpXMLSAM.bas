Attribute VB_Name = "modImpXMLSAM"
Option Compare Database
Option Explicit
Dim FN As Integer
Dim Nome_Arquivo As String
Dim strPath As String
Public XMLNumNF As String
Public XMLSerie As String
Public XMLCNPJEmi As String
Public XMLChave As String
Public XMLValNF As Double
Public XMLBCICMS As Double
Public XMLPerICMS As Double
Public XMLValICMS As Double
Public XMLDTEmi As Date
Public XMLdhSaiEnt As Date
Public XMLinfAdFisco As String
Public XMLvNF As Double
Public XMLCNPJRem As String
Public XMLCFOP As String
Public txtNode As String
Public i As Long
Public txtContaUFEmit As Long
Public txtIdentifCNPJ As Boolean

Function PegaNODXML(DocXML As String)
Dim XMLdoc As New MSXML2.DOMDocument60
Dim objNodeList As IXMLDOMNodeList
Dim i As Integer
Dim cCT
Dim CNPJ

XMLdoc.async = False
XMLdoc.Load DocXML


If (XMLdoc.parseError.errorCode <> 0) Then
    Dim myErr
    Set myErr = XMLdoc.parseError
    MsgBox ("You have error " & myErr.reason)
Else
    Set objNodeList = XMLdoc.getElementsByTagName("ide")
    cCT = objNodeList.Item(i).text
    Set objNodeList = XMLdoc.getElementsByTagName("emit")
    CNPJ = objNodeList.Item(i).text
End If

End Function

'' #DUVIDA - QUAL O OBJETIVO ?
'' #ENTENDIMENTO_01 - CARREGAR TODOS OS ITENS DA COMPRA
'' #AILTON - qryInsertCompraItens
Function PegaTagProdXML(DocXML As String, IDCompraNF As Double, IDForn As Long)
On Error GoTo Err_PegaTagProdXML

'On Error Resume Next
Dim i As Integer
Dim sBn As String
Dim qtdProd As String


Dim rsIPI As Recordset
Dim rsICMS As Recordset

Dim db As Database
Dim rsProd As Recordset
Dim qryProd As QueryDef

Dim nitem As Integer
Dim p As Integer
Dim IDCFOP As Long
Dim OutICMS As String
Dim OutIPI As String
Dim RatIPI As Double

Dim xProd As String
Dim cProd As String
Dim x
Dim CFOP As String
Dim NCM As String
Dim uCom As String
Dim qCom As String
Dim vUnCom As String
Dim vProd As String
Dim VFrete As String
Dim VDesc As String
Dim vOutro As String
Dim VICMS As String
Dim CST As String
Dim vBC As Double
Dim pRedBC  As Double
Dim pICMS  As Double
Dim pMVAST  As Double
Dim vBCST  As Double
Dim pICMSST  As Double
Dim vICMSST  As Double
Dim cEnq As String
Dim vBC_IPI As Double
Dim pIPI As Double
Dim VIPI As Double

Dim txID
Dim txGrade
Dim txCod
Dim TxDesc


Dim Doc As DOMDocument60
Dim XMLdoc As Object
Set Doc = New DOMDocument60

Set XMLdoc = CreateObject("Microsoft.XMLDOM")
XMLdoc.async = False
XMLdoc.Load (DocXML)


'' -- sBn / Qual o objetivo dessa variavel ?
qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").Length 'Contando quantos itens tem o nó det (detalhes)

p = 0
'Item = 0

Set db = CurrentDb

If (XMLdoc.parseError.errorCode <> 0) Then
    Dim myErr
    Set myErr = XMLdoc.parseError
    MsgBox ("You have error " & myErr.reason)
Else
    Dim str As String
    str = ""
    
    'Set objNodeList = XMLdoc.getElementsByTagName("prod")
    '' #AILTON - INICIO DO LOOP
    For i = 0 To qtdProd - 1 'Varrendo todos os itens
        
        'Item = Item + 1
        nitem = CStr(XMLdoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)
        cProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/cProd").Item(0).text)
        xProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/xProd").Item(0).text)
        
        Dim parts() As String
        Dim NumAtual As String
   
   
        If Not IsNull(xProd) Then
            parts = Split((xProd), " - ")
            NumAtual = parts(LBound(parts))
            If Forms!frmCompraNF_ImpXML!Finalidade = 2 Or Forms!frmCompraNF_ImpXML!Finalidade = 3 Or Forms!frmCompraNF_ImpXML!Finalidade = 4 Or Forms!frmCompraNF_ImpXML!Finalidade = 5 Or Forms!frmCompraNF_ImpXML!Finalidade = 6 Or Forms!frmCompraNF_ImpXML!Finalidade = 7 Then
                NumAtual = cProd
            End If
            If UBound(parts) = 0 And (Forms!frmCompraNF_ImpXML!Finalidade <> 2 And Forms!frmCompraNF_ImpXML!Finalidade <> 3 And Forms!frmCompraNF_ImpXML!Finalidade <> 4 And Forms!frmCompraNF_ImpXML!Finalidade <> 5 And Forms!frmCompraNF_ImpXML!Finalidade <> 6 And Forms!frmCompraNF_ImpXML!Finalidade <> 7) Then
                x = ""
            Else
                
                If Forms!frmCompraNF_ImpXML!Finalidade = 2 Then
                    '' #DUVIDA - QUAL O OBJETIVO ?
                    '' #ENTENDIMENTO_01 - CARREGAR TODOS OS ITENS DA COMPRA
                    Set qryProd = db.QueryDefs("qryCompraNF_ImpXML_ProdConsumo")
                    qryProd.Parameters(0) = NumAtual
                    qryProd.Parameters(1) = IDForn
                    
                Else
                    '' #DUVIDA - QUAL O OBJETIVO ?
                    '' #ENTENDIMENTO_01 - CARREGAR TODOS OS ITENS DA COMPRA
                    Set qryProd = db.QueryDefs("qryCompraNF_ImpXML_Prod")
                    qryProd.Parameters(0) = STRPontos(NumAtual)
                    'MsgBox STRPontos(NumAtual)
                    
                End If
                
                Set rsProd = qryProd.OpenRecordset(dbOpenDynaset, dbSeeChanges)
                
                If rsProd.RecordCount > 0 Then
                    txID = rsProd!ID_Prod
                    txGrade = rsProd!Codigo_Grade
                    txCod = rsProd!Cod_Prod
                    TxDesc = rsProd!Descr_Prod
                    
                End If
                
                rsProd.Close
            End If
        End If
       
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 - CARREGAR TODOS OS ITENS DA COMPRA
        NCM = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/NCM").Item(0).text)
        CFOP = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/CFOP").Item(0).text)
        
        
        '' #Ailton - qryUpdateFinalidade
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 - CARREGAR TODOS OS ITENS DA COMPRA
        If Forms!frmCompraNF_ImpXML!Finalidade = 2 Then
            If CFOP = "5405" Then
                CFOP = "1.407"
                IDCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.407' and [Fil_NatOper]='" & Forms!frmCompraNF_ImpXML.FiltroFil & "'")
            ElseIf CFOP = "6403" Then
                CFOP = "2.407"
                IDCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.407' and [Fil_NatOper]='" & Forms!frmCompraNF_ImpXML.FiltroFil & "'")
            Else
                CFOP = Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(1)
                IDCFOP = Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(0)
            End If
        Else
            CFOP = Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(1)
            IDCFOP = Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(0)
        End If
        
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 - CARREGAR TODOS OS ITENS DA COMPRA
        uCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/uCom").Item(0).text)
        qCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/qCom").Item(0).text)
        
        '#05_XML_IPI
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 -
        Set rsIPI = db.OpenRecordset("SELECT [05_XML_IPI].nID, [05_XML_IPI].cEnq, [05_XML_IPI].CST, [05_XML_IPI].vBC, [05_XML_IPI].pIPI, [05_XML_IPI].vIPI FROM 05_XML_IPI WHERE ([05_XML_IPI].nID=" & nitem & ") AND ([05_XML_IPI].vIPI Is Not Null) ;")
        
        If rsIPI.RecordCount > 0 And Forms!frmCompraNF_ImpXML!Finalidade = 2 Then
            RatIPI = (Replace(rsIPI!VIPI, ".", ",") / 100) / Replace(qCom, ".", ",")
            vUnCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vUnCom").Item(0).text)
            vProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vProd").Item(0).text)
            vProd = (Val(vProd)) + Val(Replace(rsIPI!VIPI, ",", "."))
            vUnCom = vProd / Replace(qCom, ".", ",")
            vProd = Replace(vProd, ",", ".")
            vUnCom = Replace(vUnCom, ",", ".")
        Else
            vUnCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vUnCom").Item(0).text)
            vProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vProd").Item(0).text)
            vUnCom = Replace(vUnCom, ",", ".")
            vProd = Replace(vProd, ",", ".")
        End If
        
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vFrete").Length > 0 Then
            VFrete = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vFrete").Item(0).text)
            VFrete = Replace(VFrete, ",", ".")
        Else
            VFrete = 0
        End If
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vDesc").Length > 0 Then
            VDesc = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vDesc").Item(0).text)
            VDesc = Replace(VDesc, ",", ".")
        Else
            VDesc = 0
        End If
        
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vOutro").Length > 0 Then
            vOutro = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vOutro").Item(0).text)
            vOutro = Replace(vOutro, ",", ".")
        Else
            vOutro = 0
        End If
        
        
        
        
        '' #05_XML_ICMS
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 -
        Set rsICMS = db.OpenRecordset("SELECT [05_XML_ICMS].nID, [05_XML_ICMS].orig, [05_XML_ICMS].CST, [05_XML_ICMS].modBC, [05_XML_ICMS].vBC, [05_XML_ICMS].pICMS, [05_XML_ICMS].vICMS, [05_XML_ICMS].pRedBC, [05_XML_ICMS].vBCSTRet, [05_XML_ICMS].vICMSSTRet, [05_XML_ICMS].modBCST, [05_XML_ICMS].vBCST, [05_XML_ICMS].pICMSST, [05_XML_ICMS].vICMSST, [05_XML_ICMS].[pMVAST] FROM 05_XML_ICMS WHERE ([05_XML_ICMS].nID=" & nitem & ") AND ([05_XML_ICMS].vICMS Is Not Null) ;")
        If rsICMS.RecordCount > 0 Then
        '' #05_XML_ICMS_Orig
        '' #05_XML_ICMS_CST
        
            CST = rsICMS!Orig & rsICMS!CST
            If Forms!frmCompraNF_ImpXML!Finalidade = 2 Then
                vBC = 0
            Else
                vBC = rsICMS!vBC
            End If
        End If
        
        'On Error Resume Next
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 -
        If Forms!frmCompraNF_ImpXML!Finalidade = 2 Then
            pRedBC = 0
            pICMS = 0
            VICMS = 0
            pMVAST = 0
            vBCST = 0
            pICMSST = 0
            vICMSST = 0
            OutICMS = Replace(Val(Replace(vProd, ".", ",")) + Val(Nz(Replace(vOutro, ".", ","))), ",", ".")
            OutIPI = Replace(Val(Replace(vProd, ".", ",")) + Val(Nz(Replace(vOutro, ".", ","))), ".", ",")
        Else
            OutICMS = 0
            OutIPI = 0
            If rsICMS!pRedBC = 0 Then
                pRedBC = 0
            ElseIf IsNull(rsICMS!pRedBC) Or rsICMS!pRedBC = "" Then
                pRedBC = 0
            Else
                pRedBC = rsICMS!pRedBC
            End If
            If rsICMS!pICMS = 0 Then
                pICMS = 0
            Else
                pICMS = rsICMS!pICMS
            End If
            
            '' #05_XML_ICMS_CST_VICMS
            If rsICMS!VICMS = 0 Then
                VICMS = 0
            Else
                VICMS = rsICMS!VICMS
            End If
            If rsICMS!pMVAST = 0 Then
                pMVAST = 0
            ElseIf IsNull(rsICMS!pMVAST) Or rsICMS!pMVAST = "" Then
                pMVAST = 0
            Else
                pMVAST = rsICMS!pMVAST
            End If
            If rsICMS!vBCST Then
                vBCST = 0
            ElseIf IsNull(rsICMS!vBCST) Or rsICMS!vBCST = "" Then
                vBCST = 0
            Else
                vBCST = rsICMS!vBCST
            End If
            If rsICMS!pICMSST = 0 Then
                pICMSST = 0
            ElseIf IsNull(rsICMS!pICMSST) Or rsICMS!pICMSST = "" Then
                pICMSST = 0
            Else
                pICMSST = rsICMS!pICMSST
            End If
            If rsICMS!vICMSST = 0 Then
                vICMSST = 0
            ElseIf IsNull(rsICMS!vICMSST) Or rsICMS!vICMSST = "" Then
                vICMSST = 0
            Else
                vICMSST = rsICMS!vICMSST
            End If
        End If
        
        'IPI
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 -
        If rsIPI.RecordCount > 0 Then
            cEnq = rsIPI!cEnq
            If Forms!frmCompraNF_ImpXML!Finalidade = 2 Then
                vBC_IPI = 0
                pIPI = 0
                VIPI = 0
            Else
                If Nz(rsIPI!pIPI, 0) = 0 Then
                    vBC_IPI = 0
                Else
                    vBC_IPI = rsIPI!vBC
                End If
                pIPI = Nz(rsIPI!pIPI, 0)
                VIPI = Nz(rsIPI!VIPI, 0)
            End If
        Else
            cEnq = "999"
            vBC_IPI = 0
            pIPI = 0
            VIPI = 0
        End If
        rsIPI.Close
       
        '' #AILTON - VALIDAR
       
        '' #AILTON - ITENS DA COMPRANF ( CONSULTA NFE )
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 - CADASTRO DE ITENS DA COMPRA SEPARADA DA TELA DE CADASTRO
        CurrentDb.Execute "INSERT INTO tblCompraNFItem ( ID_CompraNF_CompraNFItem, Item_CompraNFItem, ID_Prod_CompraNFItem, ID_Grade_CompraNFItem, QtdFat_CompraNFItem, VUnt_CompraNFItem, " _
        & "TxDesc_CompraNFItem , VUntDesc_CompraNFItem, ICMS_CompraNFItem, ISS_CompraNFItem, IPI_CompraNFItem, CFOP_CompraNFItem, ST_CompraNFItem, " _
        & "FlagEst_CompraNFItem , TxMLSubsTrib_CompraNFItem, " _
        & "Num_CompraNFItem , VTot_CompraNFItem, BaseCalcICMS_CompraNFItem, DebICMS_CompraNFItem , IseICMS_CompraNFItem, " _
        & "BaseCalcIPI_CompraNFItem, DebIPI_CompraNFItem, IseIPI_CompraNFItem,  " _
        & "BaseCalcICMSSubsTrib_CompraNFItem, VTotICMSSubsTrib_compranfitem, VTotFrete_CompraNFItem," _
        & "VTotDesc_CompraNFItem , VTotBaseCalcICMS_CompraNFItem, ID_NatOp_CompraNFItem, VTotPIS_CompraNFItem, STPIS_CompraNFItem," _
        & "PIS_CompraNFItem , VTotBaseCalcPIS_CompraNFItem, VTotCOFINS_CompraNFItem, STCOFINS_CompraNFItem, COFINS_CompraNFItem," _
        & "VTotBaseCalcCOFINS_CompraNFItem, STIPI_CompraNFItem, VTotIseICMS_CompraNFItem, SNCredICMS_CompraNFItem, " _
        & "VTotSNCredICMS_CompraNFItem, VTotSeg_CompraNFItem,  VTotOutICMS_CompraNFItem, OutIPI_CompraNFItem, VTotOutDesp_CompraNFItem, Almox_CompraNFItem ) " _
        & "SELECT " & IDCompraNF & "," & nitem & "," & txID & "," & txGrade & "," & qCom & "," & vUnCom & "," _
        & 0 & "," & 0 & "," & pICMS & "," & 0 & "," & Replace(pIPI, ",", ".") & ",'" & CFOP & "','" & 1 & Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(4) & "'," _
        & -1 & "," & pMVAST & "," _
        & 0 & "," & vProd & "," & Replace((100 - Replace(pRedBC, ".", ",")), ",", ".") & "," & Replace(VICMS, ",", ".") & "," & 0 & "," _
        & Replace(vBC_IPI, ",", ".") & "," & Replace(VIPI, ",", ".") & "," & 0 & "," _
        & vBCST & "," & vICMSST & "," & Replace(VFrete, ",", ".") & "," _
        & Replace(VDesc, ",", ".") & "," & Replace(vBC, ",", ".") & "," & IDCFOP & "," & 0 & ",'" & Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(6) & "'," _
        & 0 & "," & 0 & "," & 0 & ",'" & Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(6) & "'," & 0 & "," _
        & 0 & ",'" & Forms!frmCompraNF_ImpXML.FiltroCFOP.Column(5) & "'," & 0 & "," & 0 & "," _
        & 0 & "," & 0 & "," & Replace(OutICMS, ",", ".") & "," & Replace(OutIPI, ",", ".") & "," & Replace(vOutro, ",", ".") & "," & Forms!frmCompraNF_ImpXML!FiltroAlmox
        
    Next

End If

Exit_PegaTagProdXML:
    Exit Function
Err_PegaTagProdXML:
    MsgBox Error$
    Resume Exit_PegaTagProdXML
End Function

'' #DUVIDA - QUAL O OBJETIVO ?
'' #ENTENDIMENTO - CARREGAR DADOS DO ARQUIVO
Function LerXML(Arquivo As String)
    Dim objXML As MSXML2.DOMDocument60
    Set objXML = New MSXML2.DOMDocument60

    '' #AILTON - INICIO DE VARIAVEIS
    XMLNumNF = ""
    XMLCNPJEmi = ""
    XMLSerie = ""
    XMLChave = ""
    XMLValNF = 0
    XMLBCICMS = 0
    XMLPerICMS = 0
    XMLValICMS = 0
    XMLDTEmi = "00:00:00"
    XMLdhSaiEnt = "00:00:00"
    XMLinfAdFisco = ""
    XMLvNF = 0
    XMLCNPJRem = ""
    XMLCFOP = ""

    objXML.validateOnParse = False
    If objXML.Load(Arquivo) Then ' Verifico se carregou o XML
    
        '' #AILTON
        If Forms!frmCompraNF_ImpXML!Finalidade = 0 Or Forms!frmCompraNF_ImpXML!Finalidade = 6 Or Forms!frmCompraNF_ImpXML!Finalidade = 7 Then
            '' #DUVIDA - QUAL O OBJETIVO ?
            '' #ENTENDIMENTO_01 - CARREGAR CAMPOS DA TELA - Form_frmCompraNF_ImpXML
            LerNodesRem objXML.ChildNodes
        End If
        
        '' #DUVIDA - QUAL O OBJETIVO ?
        '' #ENTENDIMENTO_01 - CARREGAR CABEÇALHO DA COMPRA
        LerNodes objXML.ChildNodes ' Se carregou, leio os Nodes
        
    Else
    
        MsgBox "XML Não foi lido", vbCritical 'Senão, aviso que deu pau
        
    End If
    
End Function

'' #DUVIDA - QUAL O OBJETIVO ?
'' #ENTENDIMENTO_01 - CARREGAR CABEÇALHO DA COMPRA
Function LerNodes(ByRef Nodes As IXMLDOMNodeList)
   Dim objNode As IXMLDOMNode

   For Each objNode In Nodes ' Passo por todos os nodes
      If objNode.NodeType = NODE_TEXT Then
      
        If XMLNumNF = "" Then
            If Forms!frmCompraNF_ImpXML!Finalidade = 0 Then
                If objNode.ParentNode.nodeName = "nCT" Then
                    XMLNumNF = objNode.NodeValue
                End If
            Else
                If objNode.ParentNode.nodeName = "nNF" Then
                    XMLNumNF = objNode.NodeValue
                End If
            End If
        End If
        If XMLSerie = "" Then
            If Forms!frmCompraNF_ImpXML!Finalidade = 0 Then
                If objNode.ParentNode.nodeName = "serie" Then
                    XMLSerie = objNode.NodeValue
                End If
            Else
                If objNode.ParentNode.nodeName = "serie" Then
                    XMLSerie = objNode.NodeValue
                End If
            End If
        End If
        
        If XMLinfAdFisco = "" Then
            If objNode.ParentNode.nodeName = "infAdFisco" Then
                XMLinfAdFisco = objNode.NodeValue
            End If
        End If
        
        If XMLCFOP = "" Then
            If objNode.ParentNode.nodeName = "CFOP" Then
                XMLCFOP = objNode.NodeValue
            End If
        End If
        
        If XMLvNF = 0 Then
            If objNode.ParentNode.nodeName = "vNF" Then
                XMLvNF = objNode.NodeValue / 100
            End If
        End If
        
        
        If XMLCNPJEmi = "" Then
            If objNode.ParentNode.nodeName = "CNPJ" Then
                'If objNode.NodeValue = STRPontos(DLookup("[CNPJ_EMPRESA]", "tblEmpresa")) Then
                If objNode.NodeValue = STRPontos(DLookup("[CNPJ_Empresa]", "tblEmpresa", "[ID_Empresa]='" & Forms!frmCompraNF_ImpXML!FiltroFil & "'")) Then
                Else
                    XMLCNPJEmi = objNode.NodeValue
                End If
            End If
        End If
        If XMLValNF = 0 Then
            If objNode.ParentNode.nodeName = "vTPrest" Then
                XMLValNF = objNode.NodeValue / 100
            End If
        End If
        If XMLBCICMS = 0 Then
            If objNode.ParentNode.nodeName = "vBC" Then
                XMLBCICMS = objNode.NodeValue / 100
            End If
        End If
        If XMLPerICMS = 0 Then
            If objNode.ParentNode.nodeName = "pICMS" Then
                XMLPerICMS = objNode.NodeValue / 100
            End If
        End If
        
        If XMLValICMS = 0 Then
            If objNode.ParentNode.nodeName = "vICMS" Then
                XMLValICMS = objNode.NodeValue / 100
            End If
        End If
        
        If XMLDTEmi = "00:00:00" Then
            If Forms!frmCompraNF_ImpXML!Finalidade = 0 Then
                If objNode.ParentNode.nodeName = "dhEmi" Then
                    XMLDTEmi = CDate(Replace(Mid(objNode.NodeValue, 1, 10), "-", "/"))
                End If
            Else
                If objNode.ParentNode.nodeName = "dhEmi" Then
                    XMLDTEmi = CDate(Replace(Mid(objNode.NodeValue, 1, 10), "-", "/"))
                End If
            End If
        End If
        
        
        ''HoraEntd_CompraNF
        If XMLdhSaiEnt = "00:00:00" Then
            If Forms!frmCompraNF_ImpXML!Finalidade = 4 Then
                If objNode.ParentNode.nodeName = "dhSaiEnt" Then
                    XMLdhSaiEnt = (Replace(Mid(objNode.NodeValue, 12, 8), "-", "/"))
                End If
            End If
        End If

        If XMLChave = "" Then
            If Forms!frmCompraNF_ImpXML!Finalidade = 0 Then
                If objNode.ParentNode.nodeName = "chCTe" Then
                    XMLChave = objNode.NodeValue
                End If
            Else
                If objNode.ParentNode.nodeName = "chNFe" Then
                    XMLChave = objNode.NodeValue
                End If
            
            End If
        End If
        
        
        If Forms!frmCompraNF_ImpXML!Finalidade = 1 Then
            If objNode.ParentNode.nodeName = "det nItem" Then
            MsgBox "teste"
            End If
        End If
        
        
      End If
      
     
        If objNode.HasChildNodes Then 'Verifico se é ChildNode, se for leio o próximo.
            LerNodes objNode.ChildNodes
        End If
   Next objNode

End Function

'' #DUVIDA - QUAL O OBJETIVO ?
'' #ENTENDIMENTO_01 - CARREGAR CAMPOS DA TELA - Form_frmCompraNF_ImpXML
Function LerNodesRem(ByRef Nodes As IXMLDOMNodeList)
Dim objNode As IXMLDOMNode

    If i > 0 Then  'CNPJ Remetente
    Else
        i = 0
    End If
    
    If txtContaUFEmit > 0 Then  'UF Emitente
    Else
        txtContaUFEmit = 0
    End If

    For Each objNode In Nodes ' Passo por todos os nodes
    
        '' #PONTO - CNPJ
        If objNode.ParentNode.nodeName = "CNPJ" Then
             i = i + 1
        End If
        
        '' #PONTO - UF
        If objNode.ParentNode.nodeName = "UF" Then
            txtContaUFEmit = txtContaUFEmit + 1
        End If
    
        '' #PONTO - txtContaUFEmit
        If txtContaUFEmit = 1 Then
          If objNode.NodeType = NODE_TEXT Then '
             If objNode.ParentNode.nodeName = "UF" Then
                Forms!frmCompraNF_ImpXML!txtUFEmit = objNode.NodeValue
             End If
          End If
        End If
     
        '' #PONTO - FINALIDADE
        If Forms!frmCompraNF_ImpXML!Finalidade = 0 Then
            
            '' #PONTO - txtIdentifCNPJ
             If i = 2 And txtIdentifCNPJ = False Then
                If objNode.NodeType = NODE_TEXT Then '
                   If objNode.ParentNode.nodeName = "CNPJ" Then
                      Forms!frmCompraNF_ImpXML!txtCNPJRem = objNode.NodeValue
                      Forms!frmCompraNF_ImpXML!FiltroFil = DLookup("[ID_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'")
                      
                      If Not IsNull(Forms!frmCompraNF_ImpXML!FiltroFil) Or Forms!frmCompraNF_ImpXML!FiltroFil <> "" Then
                          If Forms!frmCompraNF_ImpXML!FiltroFil = "PSP" Then
                              Forms!frmCompraNF_ImpXML!FiltroAlmox = 6
                              Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                              
                          ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PSC" Then
                              Forms!frmCompraNF_ImpXML!FiltroAlmox = 12
                              Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                              
                          ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PES" Then
                              Forms!frmCompraNF_ImpXML!FiltroAlmox = 1
                              Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                          End If
                          
                          txtIdentifCNPJ = True
                          
                      End If
                   End If
                End If
                
              '' #PONTO - txtIdentifCNPJ
              ElseIf i = 3 And txtIdentifCNPJ = False Then
              
                  If objNode.NodeType = NODE_TEXT Then '
                      If objNode.ParentNode.nodeName = "CNPJ" Then
                          Forms!frmCompraNF_ImpXML!txtCNPJRem = objNode.NodeValue
                          Forms!frmCompraNF_ImpXML!FiltroFil = DLookup("[ID_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'")
                          
                          If Not IsNull(Forms!frmCompraNF_ImpXML!FiltroFil) Or Forms!frmCompraNF_ImpXML!FiltroFil <> "" Then
                          
                              If Forms!frmCompraNF_ImpXML!FiltroFil = "PSP" Then
                                  Forms!frmCompraNF_ImpXML!FiltroAlmox = 6
                                  Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                  
                              ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PSC" Then
                                  Forms!frmCompraNF_ImpXML!FiltroAlmox = 12
                                  Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                  
                              ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PES" Then
                                  Forms!frmCompraNF_ImpXML!FiltroAlmox = 1
                                  Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                              End If
                              
                              txtIdentifCNPJ = True
                              
                          End If
                       End If
                    End If
                    
              End If
         
         '' #PONTO - FINALIDADE
         ElseIf Forms!frmCompraNF_ImpXML!Finalidade = 6 Or Forms!frmCompraNF_ImpXML!Finalidade = 7 Then
         
                 '' #PONTO - txtIdentifCNPJ
                 If i = 2 And txtIdentifCNPJ = False Then
                 
                   If objNode.NodeType = NODE_TEXT Then
                      If objNode.ParentNode.nodeName = "CNPJ" Then
                         Forms!frmCompraNF_ImpXML!txtCNPJRem = objNode.NodeValue
                         Forms!frmCompraNF_ImpXML!FiltroFil = DLookup("[ID_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'")
                         
                         If Not IsNull(Forms!frmCompraNF_ImpXML!FiltroFil) Or Forms!frmCompraNF_ImpXML!FiltroFil <> "" Then
                         
                             '' #PONTO - FINALIDADE
                             If Forms!frmCompraNF_ImpXML!Finalidade = 6 Then
                                 'Almoxarifado: Puxa 97 Empresa for PES / 98 Empresa for PSP / 99 Empresa for PSC
                         
                                 If Forms!frmCompraNF_ImpXML!FiltroFil = "PSP" Then
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox = 98
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                     
                                 ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PSC" Then
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox = 99
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                     
                                 ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PES" Then
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox = 97
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                 End If
                                 
                             '' #PONTO - FINALIDADE
                             ElseIf Forms!frmCompraNF_ImpXML!Finalidade = 7 Then
                                 'Almoxarifado: Puxa 1 Empresa for PES / 2 Empresa for PSP / 12 Empresa for PSC
                         
                                 If Forms!frmCompraNF_ImpXML!FiltroFil = "PSP" Then
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox = 2
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                     
                                 ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PSC" Then
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox = 12
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                     
                                 ElseIf Forms!frmCompraNF_ImpXML!FiltroFil = "PES" Then
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox = 1
                                     Forms!frmCompraNF_ImpXML!FiltroAlmox.Requery
                                 End If
                                 
                             End If
                             
                             txtIdentifCNPJ = True
                             
                         End If
                      End If
                   End If
                 
                 '' #PONTO - txtIdentifCNPJ
                 ElseIf i = 3 And txtIdentifCNPJ = False Then
                 
                     If objNode.NodeType = NODE_TEXT Then '
                         If objNode.ParentNode.nodeName = "CNPJ" Then
                             Forms!frmCompraNF_ImpXML!txtCNPJDest = objNode.NodeValue
                             If Not IsNull(Forms!frmCompraNF_ImpXML!FiltroFil) Or Forms!frmCompraNF_ImpXML!FiltroFil <> "" Then
                                 txtIdentifCNPJ = True
                             End If
                          End If
                     End If
                     
                 End If
        
        End If
        
        If objNode.HasChildNodes Then 'Verifico se é ChildNode, se for leio o próximo.
            LerNodesRem objNode.ChildNodes
        End If
    
    Next objNode
    
End Function

Public Function Imp_Prod_ICMS(sFile As String)
On Error Resume Next
Dim i As Integer
Dim sBn As String
Dim qtdProd As String

Dim nitem As String
Dim Orig As String
Dim CST As String
Dim modBC As String
Dim pICMS As String
Dim vBC As String
Dim VICMS As String
Dim pRedBC As String
Dim modBCST As String
Dim pMVAST As String
Dim pRedBCST As String
Dim vBCST As String
Dim pICMSST As String
Dim vICMSST As String
                    

Dim Doc As DOMDocument60
Dim XMLdoc As Object
Set Doc = New DOMDocument60

Set XMLdoc = CreateObject("Microsoft.XMLDOM")
XMLdoc.async = False
XMLdoc.Load (sFile)

qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").Length 'Contando quantos itens tem o nó det (detalhes)
    
If (XMLdoc.parseError.errorCode <> 0) Then
    Dim myErr
    Set myErr = XMLdoc.parseError
    MsgBox ("You have error " & myErr.reason)
Else
    Dim str As String
    str = ""
    For i = 0 To qtdProd - 1 'Varrendo todos os itens
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS").Length > 0 Then
            Orig = Null
            CST = Null
            modBC = Null
            vBC = 0
            pICMS = 0
            VICMS = 0
            modBCST = 0
            pMVAST = 0
            pRedBCST = 0
            vBCST = 0
            pICMSST = 0
            vICMSST = 0

            nitem = CStr(XMLdoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)

            'ICMS 00
            If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Item(0).getElementsByTagName("CST").Item(0).text
                modBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Item(0).getElementsByTagName("modBC").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Item(0).getElementsByTagName("vBC").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Item(0).getElementsByTagName("pICMS").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS00").Item(0).getElementsByTagName("vICMS").Item(0).text
            'ICMS 10
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("CST").Item(0).text
                modBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("modBC").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("vBC").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("pICMS").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("vICMS").Item(0).text
                modBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("modBCST").Item(0).text
                pMVAST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("pMVAST").Item(0).text
                pRedBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("pRedBCST").Item(0).text
                vBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("vBCST").Item(0).text
                pICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("pICMSST").Item(0).text
                vICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS10").Item(0).getElementsByTagName("vICMSST").Item(0).text
            'ICMS 20
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("CST").Item(0).text
                modBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("modBC").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("vBC").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("pICMS").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("vICMS").Item(0).text
                pRedBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS20").Item(0).getElementsByTagName("pRedBC").Item(0).text
            'ICMS 30
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("CST").Item(0).text
                modBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("modBCST").Item(0).text
                pMVAST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("pMVAST").Item(0).text
                pRedBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("pRedBCST").Item(0).text
                vBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("vBCST").Item(0).text
                pICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("pICMSST").Item(0).text
                vICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS30").Item(0).getElementsByTagName("vICMSST").Item(0).text
            'ICMS 40, 41, 50
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS40").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS40").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS40").Item(0).getElementsByTagName("CST").Item(0).text
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS41").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS41").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS41").Item(0).getElementsByTagName("CST").Item(0).text
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS50").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS50").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS50").Item(0).getElementsByTagName("CST").Item(0).text
            'ICMS 51
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("CST").Item(0).text
                modBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("modBC").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("vBC").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("pICMS").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("vICMS").Item(0).text
                pRedBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS51").Item(0).getElementsByTagName("pRedBC").Item(0).text
            'ICMS 60
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS60").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS60").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS60").Item(0).getElementsByTagName("CST").Item(0).text
            'ICMS 70
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("CST").Item(0).text
                modBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("modBC").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("vBC").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("pICMS").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("vICMS").Item(0).text
                pRedBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("pRedBC").Item(0).text
                modBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("modBCST").Item(0).text
                pMVAST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("pMVAST").Item(0).text
                pRedBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("pRedBCST").Item(0).text
                vBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("vBCST").Item(0).text
                pICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("pICMSST").Item(0).text
                vICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS70").Item(0).getElementsByTagName("vICMSST").Item(0).text
            'ICMS 90
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("CST").Item(0).text
                modBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("modBC").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("vBC").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("pICMS").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("vICMS").Item(0).text
                pRedBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("pRedBC").Item(0).text
                modBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("modBCST").Item(0).text
                pMVAST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("pMVAST").Item(0).text
                pRedBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("pRedBCST").Item(0).text
                vBCST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("vBCST").Item(0).text
                pICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("pICMSST").Item(0).text
                vICMSST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMS90").Item(0).getElementsByTagName("vICMSST").Item(0).text
            'ICMSSN 102
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN102").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN102").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN102").Item(0).getElementsByTagName("CSOSN").Item(0).text
            'ICMSSN 500
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN500").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN500").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN500").Item(0).getElementsByTagName("CSOSN").Item(0).text
            'ICMSSN 101
            ElseIf XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN101").Length > 0 Then
                Orig = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN101").Item(0).getElementsByTagName("orig").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN101").Item(0).getElementsByTagName("CSOSN").Item(0).text
                pICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN101").Item(0).getElementsByTagName("pCredSN").Item(0).text
                VICMS = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/ICMS/ICMSSN101").Item(0).getElementsByTagName("vCredICMSSN").Item(0).text
            End If
        End If
        
        '' #05_XML_ICMS
        CurrentDb.Execute "INSERT INTO 05_XML_ICMS ( id, nID, orig, CST, modBC, vBC, pICMS, vICMS, pRedBC, " _
        & "modBCST, vBCST, pICMSST, vICMSST, pMVAST )" _
        & "SELECT '" & nitem & "' AS Expr1, '" & nitem & "' AS Expr2, '" & Orig & "' AS Expr3, '" & CST & "' AS Expr4, '" & modBC & "' AS Expr5, " _
        & "" & vBC & " AS Expr6, " & pICMS & " AS Expr7, " & VICMS & " AS Expr8, " & IIf(pRedBC = "", 0, pRedBC) & " AS Expr9, " _
        & "" & IIf(modBCST = "", 0, modBCST) & " AS Expr12, " & IIf(vBCST = "", 0, vBCST) & " AS Expr13, " & IIf(pICMSST = "", 0, pICMSST) & " AS Expr14, " & IIf(vICMSST = "", 0, vICMSST) & " AS Expr15, " & IIf(pMVAST = "", 0, pMVAST) & " AS Expr16 "

    Next i

End If

End Function



Public Function Imp_Prod_IPI(sFile As String)
On Error Resume Next
Dim i As Integer
Dim sBn As String
Dim qtdProd As String

Dim nitem As String
Dim cEnq As String
Dim CST As String
Dim pIPI As String
Dim vBC As String
Dim VIPI As String
                    

Dim Doc As DOMDocument60
Dim XMLdoc As Object
Set Doc = New DOMDocument60

Set XMLdoc = CreateObject("Microsoft.XMLDOM")
XMLdoc.async = False
XMLdoc.Load (sFile)


qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").Length 'Contando quantos itens tem o nó det (detalhes)
    
If (XMLdoc.parseError.errorCode <> 0) Then
    Dim myErr
    Set myErr = XMLdoc.parseError
    MsgBox ("You have error " & myErr.reason)
Else
    Dim str As String
    str = ""
    For i = 0 To qtdProd - 1 'Varrendo todos os itens
        cEnq = "999"
        CST = Null
        vBC = 0
        pIPI = 0
        VIPI = 0
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI").Length > 0 Then
            nitem = CStr(XMLdoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)
            
            If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI/IPITrib").Length > 0 Then
                cEnq = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI").Item(0).getElementsByTagName("cEnq").Item(0).text
                CST = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI/IPITrib").Item(0).getElementsByTagName("CST").Item(0).text
                vBC = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI/IPITrib").Item(0).getElementsByTagName("vBC").Item(0).text
                pIPI = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI/IPITrib").Item(0).getElementsByTagName("pIPI").Item(0).text
                VIPI = XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("imposto/IPI/IPITrib").Item(0).getElementsByTagName("vIPI").Item(0).text
            End If
        End If
        

        If pIPI <> "" Then
            CurrentDb.Execute "INSERT INTO 05_XML_IPI ( id, nID, cEnq, CST, vBC, pIPI, vIPI " _
            & " )" _
            & "SELECT '" & nitem & "' AS Expr1, '" & nitem & "' AS Expr2, '" & cEnq & "' AS Expr3, '" & CST & "' AS Expr4, " _
            & "" & vBC & " AS Expr6, " & pIPI & " AS Expr7, " & VIPI & " AS Expr8 " _
            & "" & "" _
            & ""
        End If

    Next i

End If
End Function
Public Function Imp_Prod_NF(sFile As String)
Dim gro As Recordset
Dim rsProdXML As Recordset
Dim rsProdMe As New ADODB.Recordset
Dim rsProdMeCod As New ADODB.Recordset
Dim sBn As String
Dim qtdProd As String
Dim i As Integer

Dim nitem As String
Dim xProd As String
Dim cProd As String
Dim cEAN As String
Dim cEANTrib As String
Dim CFOP As String
Dim NCM As String
Dim uCom As String
Dim qCom As String
Dim vUnCom As String
Dim uTrib As String
Dim qTrib As String
Dim vUnTrib As String
Dim VDesc As String
Dim indTot As String

Dim vProd As String
Dim VFrete As String
Dim vOutro As String

Dim xPed As String
Dim nItemPed As String

Dim VICMS As String
Dim CST As String
Dim vBC As Double
Dim pRedBC  As Double
Dim pICMS  As Double
Dim pMVAST  As Double
Dim vBCST  As Double
Dim pICMSST  As Double
Dim vICMSST  As Double
Dim cEnq As String
Dim vBC_IPI As Double
Dim pIPI As Double
Dim VIPI As Double

Dim txID
Dim txGrade
Dim txCod
Dim TxDesc


Dim db As Database
Set db = CurrentDb

Dim Doc As DOMDocument60
Dim XMLdoc As Object
Set Doc = New DOMDocument60

Set XMLdoc = CreateObject("Microsoft.XMLDOM")
XMLdoc.async = False
XMLdoc.Load (sFile)

qtdProd = XMLdoc.getElementsByTagName(sBn & "infNFe/det").Length 'Contando quantos itens tem o nó det (detalhes)

If (XMLdoc.parseError.errorCode <> 0) Then
    Dim myErr
    Set myErr = XMLdoc.parseError
    MsgBox ("You have error " & myErr.reason)
Else
    For i = 0 To qtdProd - 1 'Varrendo todos os itens
        nitem = CStr(XMLdoc.getElementsByTagName("nfeProc/NFe/infNFe/det").Item(i).Attributes(0).value)             '' Item_CompraNFItem
        cProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/cProd").Item(0).text)   '' ID_Prod_CompraNFItem
        xProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/xProd").Item(0).text)   ''
        cEAN = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/cEAN").Item(0).text)
        NCM = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/NCM").Item(0).text)
        CFOP = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/CFOP").Item(0).text)     '' CFOP_CompraNFItem
        uCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/uCom").Item(0).text)
        qCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/qCom").Item(0).text)
        vUnCom = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vUnCom").Item(0).text) '' VUnt_CompraNFItem
        vProd = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vProd").Item(0).text)
        cEANTrib = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/cEANTrib").Item(0).text)
        uTrib = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/uTrib").Item(0).text)
        qTrib = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/qTrib").Item(0).text)
        indTot = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/indTot").Item(0).text)
        
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vFrete").Length > 0 Then
            VFrete = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vFrete").Item(0).text) '' VTotFrete_CompraNFItem
        Else
            VFrete = 0
        End If
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vDesc").Length > 0 Then
            VDesc = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/vDesc").Item(0).text)   '' VTotDesc_CompraNFItem
        Else
            VDesc = 0
        End If
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/xPed").Length > 0 Then
            xPed = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/xPed").Item(0).text)
        End If
        If XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/nItemPed").Length > 0 Then
            nItemPed = CStr(XMLdoc.SelectNodes("nfeProc/NFe/infNFe/det").Item(i).SelectNodes("prod/nItemPed").Item(0).text) '' Item_CompraNFItem
        End If
        
        Set gro = db.OpenRecordset("04_XML_prod", dbOpenDynaset)
        gro.AddNew
        gro!ID = nitem
        gro!nID = nitem
        gro!cProd = cProd
        gro!cEAN = cEAN
        gro!xProd = xProd
        gro!NCM = NCM
        gro!CFOP = CFOP
        gro!uCom = uCom
        gro!qCom = qCom
        gro!vUnCom = vUnCom
        gro!vProd = vProd
        gro!cEANTrib = cEANTrib
        gro!uTrib = uTrib
        gro!qTrib = qTrib
        gro!indTot = indTot
        gro!VFrete = VFrete
        gro!VDesc = VDesc
        gro!xPed = xPed
        gro!nItemPed = nItemPed
        gro.Update
    Next i
End If
gro.Close

End Function

Public Function CadProdXMLConsumo(IDForn)
Dim rsProdXML As Recordset
Dim rsProdMe As New ADODB.Recordset
Dim rsGradeMe As New ADODB.Recordset
Dim rsEstMe As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
'Dim rsProdMeCod As New ADODB.Recordset
Dim db As Database
Dim IDProd As Long
'Dim CodProd As String

Set db = CurrentDb
Set rsProdXML = db.OpenRecordset("04_XML_prod", dbOpenDynaset)

If rsProdXML.RecordCount = 0 Then
    rsProdXML.Close
    Exit Function
End If

AbrirConexao

rsProdXML.MoveFirst
While Not rsProdXML.EOF
    rsProdMe.CursorLocation = adUseClient
    rsProdMe.CursorType = adOpenKeyset
    rsProdMe.LockType = adLockOptimistic
    'rsProdMe.Open "select * from tblProd where (FlagMatProdServ_Prod= '3') AND (tblProd.IDFornUltCompra_Prod=" & IDForn & ") AND (tblProd.CodForn_Prod='" & rsProdXML!cProd & "'); ", CNN
    rsProdMe.Open "SELECT * FROM [Cadastro de Produtos] INNER JOIN tabGradeProdutos ON [Cadastro de Produtos].Código = tabGradeProdutos.CodigoProd_Grade " _
    & "WHERE ((([Cadastro de Produtos].FlagMatProdServ_Prod)='3') AND (([Cadastro de Produtos].IDFornUltCompra_Prod)=" & IDForn & ") AND ((tabGradeProdutos.CodigoForn_Grade)='" & rsProdXML!cProd & "'));", CNN

        
    If rsProdMe.RecordCount = 0 Then
    
'        rsProdMeCod.CursorLocation = adUseClient
'        rsProdMeCod.CursorType = adOpenKeyset
'        rsProdMeCod.LockType = adLockOptimistic
'        rsProdMeCod.Open "SELECT tblProd.Cod_Prod, tblProd.FlagMatProdServ_Prod, tblProd.Cod_Prod FROM tblProd " _
'        & "WHERE (((tblProd.FlagMatProdServ_Prod) = '3') And ((tblProd.Cod_Prod) Like 'UC%')) " _
'        & "ORDER BY tblProd.Cod_Prod DESC;", CNN
'
'        If rsProdMeCod.RecordCount = 0 Then
'            CodProd = "UC" & "000001"
'        Else
'            CodProd = "UC" & Format(Mid(rsProdMeCod!Cod_Prod, 3, 6) + 1, "000000")
'        End If
'        rsProdMeCod.Close
        
        
        
        '' #AILTON - qryInsertProdutoConsumo
        rsProdMe.AddNew
        rsProdMe!FlagMatProdServ_Prod = 3
        rsProdMe!Und_Prod = Mid(rsProdXML!uCom, 1, 3)
        rsProdMe!NCM = rsProdXML!NCM
        rsProdMe!Modelo = Mid(rsProdXML!xProd, 1, 60)
        rsProdMe!Ativo = True
        rsProdMe!IDMod_Prod = 7813
        rsProdMe!CodCateg = 789
        
        rsProdMe!FlagEmLinha = 0
        rsProdMe!FlagAmostra = 0
        rsProdMe!FlagAssisTec = 0
        rsProdMe!FlagNS = 0
        rsProdMe!DisKit = 0
        rsProdMe!FlagSite = 0
        rsProdMe!FlagCompInt = 0
        rsProdMe!FlagPromo = 0
        rsProdMe!FlagBaixoGiro = 0
        
        rsProdMe!Atacado = 0
        rsProdMe!AtacadoESES = 0
        rsProdMe!AtacadoSPSP = 0
        rsProdMe!AtacadoSPFora = 0
        
        rsProdMe!IDFornUltCompra_Prod = IDForn
        rsProdMe!DTCad_Prod = Date
        rsProdMe!FlagEst_Prod = 0
        rsProdMe![Peso L] = 0
        rsProdMe![Peso B] = 0

        rsProdMe!FOB = 0
        rsProdMe!FOBD = 0

        rsProdMe!Custo = 0
        rsProdMe!CustoMedioSI = 0
        rsProdMe!CustoMedio = 0
        rsProdMe![Custo Controle] = 0
        rsProdMe!Atacado = 0
        rsProdMe!Varejo = 0
        rsProdMe!VarejoSD = 0
        rsProdMe!VarejoCD = 0
        rsProdMe!VarejoOrKit = 0
        rsProdMe!VarejoKit = 0
        rsProdMe!Funcionário = 0
        rsProdMe!Patrocinado = 0
        rsProdMe!Sócio = 0
        rsProdMe!Controle = 0
        rsProdMe!Real = True
        rsProdMe!Oficial = 0
        rsProdMe!Paralelo = 0
        rsProdMe!ValMSRP = 0
        rsProdMe!LimPerDescCom = 0
        rsProdMe.Update
        
        IDProd = rsProdMe!Código
        
        rsGradeMe.CursorLocation = adUseClient
        rsGradeMe.CursorType = adOpenKeyset
        rsGradeMe.LockType = adLockOptimistic
        rsGradeMe.Open "SELECT TOP 1 * FROM tabGradeProdutos", CNN
        
        '' #AILTON/FERNANDA - qryInsertGradeProdutos
        rsGradeMe.AddNew
        rsGradeMe!CodigoProd_Grade = IDProd
        rsGradeMe!Codigo_Grade = 1
        rsGradeMe!QtdeEst_Grade = 0
        rsGradeMe!CodigoForn_Grade = rsProdXML!cProd
        rsGradeMe!CodigoBar_Grade = IIf(rsProdXML!cEANTrib = "" Or rsProdXML!cEANTrib = "SEM GTIN", Null, rsProdXML!cEANTrib)
        rsGradeMe!Atacado_Grade = 0
        rsGradeMe!Ativo_Grade = True
        rsGradeMe.Update
        rsGradeMe.Close
        
        
        '' #AILTON - >>> PENDENTE <<<
        '' #AILTON - >>> PENDENTE <<<
        '' #AILTON - >>> PENDENTE <<<
        '' #AILTON - >>> PENDENTE <<<
        '' #AILTON - >>> PENDENTE <<<
        
        
        '' #AILTON - qrySelectEmpresa_FiltroFil
        'Cadastro de Empresa
        rsEmp.CursorLocation = adUseClient
        rsEmp.CursorType = adOpenKeyset
        rsEmp.LockType = adLockOptimistic
        rsEmp.Open "SELECT tblEmpresa.ID_Empresa FROM tblEmpresa WHERE SUBSTRING (ID_Empresa,1,1)='P'", CNN
        
        '' #AILTON/FERNANDA - qrySelectProd_Est - ( NÃO NO ARQUIVO DE BACKUP )
        '' #OBJETIVO - ???
        rsEstMe.CursorLocation = adUseClient
        rsEstMe.CursorType = adOpenKeyset
        rsEstMe.LockType = adLockOptimistic
        rsEstMe.Open "SELECT TOP 1 * FROM tblProd_Est", CNN
        
        While Not rsEmp.EOF

            rsEstMe.AddNew
            rsEstMe!CodigoProd_Est = IDProd
            rsEstMe!CodUnid_Est = rsEmp!ID_Empresa
            rsEstMe!ST_Est = "000"
            rsEstMe!ID_MLICMSSubs_Est = Null
            rsEstMe.Update
            rsEmp.MoveNext
            
        Wend
        
        rsEmp.Close
        rsEstMe.Close
        
        rsProdMe.Close
    Else
        rsProdMe.Close
    End If

    rsProdXML.MoveNext
Wend


rsProdXML.Close
End Function



