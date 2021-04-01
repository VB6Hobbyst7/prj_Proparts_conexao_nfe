VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCompraNF_ImpXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim strXMLjaImp, strXMLnovo As String
Dim DocXML As String
Dim MyFile, MyName, MyPath
Dim verificaImpXML As Boolean
Dim CountXMLImp, tipImpor, CountXMLNovo As Integer

Private Sub btnArquivo_Click()
On Error GoTo TrataErro
Dim strTemp As String
tipImpor = 0

If Me.Finalidade = 0 Then
    If Me.FiltroFil = "PSP" Then
        strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
    ElseIf Me.FiltroFil = "PES" Then
        strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
    Else
        strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
    End If
    'strTemp = AbreSelArquivo(hWnd, , "P:\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL\FILIAL\NF RECEBIDAS", "Todos os arquivos" & Chr(0) & "*.xml")
Else
    strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\ProParts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
End If

If Trim(" " & strTemp) <> "" Then
 Me.caminho = strTemp
End If
strTemp = ""
TrataErro_Exit:
    Exit Sub

TrataErro:
    MsgBox err.description
    DoCmd.Hourglass False 'Mouse Ampulheta
    Resume Next 'TrataErro_Exit
End Sub

Private Sub btnPasta_Click()
tipImpor = 1

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .Title = "Selecione uma pasta"
    
    If Me.Finalidade = 0 Then
        If Me.FiltroFil = "PSP" Then
            .InitialFileName = "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL"
        ElseIf Me.FiltroFil = "PES" Then
            .InitialFileName = "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL"
        Else
            .InitialFileName = "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL"
        End If
        '.InitialFileName = "P:\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL\FILIAL\NF RECEBIDAS"
    Else
        .InitialFileName = "\\192.168.10.200\ProParts\Administrativo\Depto Fiscal\ARQUIVOS ELETRÔNICOS - FISCAL"
    End If

    If .Show = True Then
        MyPath = .SelectedItems(1)
        Me.caminho = MyPath
    End If
End With


End Sub

Private Sub Fechar_Click()
    DoCmd.Close
End Sub


Private Sub FiltroFil_AfterUpdate()
    Me.FiltroCFOP.Requery
    Me.FiltroAlmox = Null
    Me.FiltroAlmox.Requery
End Sub

Private Sub ImportarPasta_Click()

Dim db As Database
Dim rs As Recordset

On Error GoTo ErroPasta

If txtTipoImp = "V" Then
Else
    If EstáAberto("formVendas") Then
        DoCmd.Close acForm, "formVendas"
    End If
End If

    CountXMLImp = 0
    CountXMLNovo = 0

If IsNull(Me.caminho) Then
    MsgBox "Selecione uma pasta ou arquivo", vbCritical, "Selecione"
    Exit Sub
End If

'Verificar se está importando um arquivo ou uma pasta
'Variável tipImpor altera de acordo com o botão clicado, arquivo ou pasta
If tipImpor = 0 Then
    DocXML = Me.caminho
    Call Imprimir_Click
ElseIf tipImpor = 1 Then

    MyName = Dir(Me.caminho & "\" & "*.xml")

    Do While MyName <> ""
        MyName = Me.caminho & "\" & MyName
        If MyName <> "." And MyName <> ".." Then
        
        DocXML = MyName
        Call Imprimir_Click ''#AILTON - PONTO DE ENTRADA
        
        'Verifica se determinado XML já foi importado anteriormente
        'através da variável CountXMLImp, que conta o números de XML que já estavam importados, na Imprimir_Click
        If CountXMLImp <= 0 Then
        
        'Variável booleana que verifica se a rotina de importação foi bem sucedida
        'é definida como False no início da Imprimir_Click e como True no final
            If verificaImpXML = False Then Exit Sub
            
            End If
        End If
        MyName = Dir
    Loop
        
        
End If
        'Caso o XML seja novo, mostrará o msgbox
        
        If CountXMLNovo > 0 Then
        MsgBox "Importação Concluída!" & Chr(10) & Chr(10) & "XML's importados(Nº da Nota):" & Chr(10) & Chr(10) & strXMLnovo, vbInformation, "Importação realizada!"
        End If
        
        'Caso já tenha sido importado determinado XML, mostrará o msgbox
        
        If CountXMLImp > 0 Then
        MsgBox "Os XML's abaixo já foram importados anteriormente:(Nº da Nota)" & Chr(10) & Chr(10) & strXMLjaImp, vbCritical, "Importação já existente!"
        End If
        
        strXMLnovo = ""
        strXMLjaImp = ""
        DocXML = ""

Exit Sub
ErroPasta:
    MsgBox "Erro número: " & err.Number & Chr(10) & "Descrição: " & err.description, vbCritical, "Erro"
    strXMLnovo = ""
    strXMLjaImp = ""
    CountXMLImp = 0
    CountXMLNovo = 0
    DocXML = ""
End Sub

Private Sub Imprimir_Click() '' #AILTON - COMPRANF
On Error GoTo TrataErro
Dim strSql As String

Dim rsCad As ADODB.Recordset: Set rsCad = New ADODB.Recordset
Dim rsCompraNF As ADODB.Recordset: Set rsCompraNF = New ADODB.Recordset
Dim rsCompraNFItem As ADODB.Recordset: Set rsCompraNFItem = New ADODB.Recordset

Dim IDCompraNF As Double
Dim IDCadFor As Long

verificaImpXML = False

If IsNull(Me.Finalidade) Or IsNull(Me.caminho) Then
    MsgBox "Digite finalidade e selecione arquivo!", vbCritical, "Atenção"
    Me.Finalidade.SetFocus
    Exit Sub
End If

If Me.Finalidade = 0 Or Me.Finalidade = 6 Or Me.Finalidade = 7 Then
    
    '' #AILTON - PONTO DE ENTRADA
    '' #DUVIDA - QUAL O OBJETIVO ?
    '' #ENTENDIMENTO - CARREGAR DADOS DO ARQUIVO
    x = LerXML(DocXML)
    
Else

    
    If IsNull(Me.FiltroCFOP) Then
        MsgBox "Digite CFOP!", vbCritical, "Atenção"
        Me.FiltroCFOP.SetFocus
        Exit Sub
    End If
    
    '' #AILTON
    '' qrySelectEmpresa_FiltroFil
    If IsNull(Me.FiltroFil) Then
        MsgBox "Selecione a Empresa!", vbCritical, "Atenção"
        Me.FiltroFil.SetFocus
        Exit Sub
    End If
    
    If IsNull(Me.FiltroAlmox) Then
        MsgBox "Selecione almoxarifado!", vbCritical, "Atenção"
        Me.FiltroAlmox.SetFocus
        Exit Sub
    End If
    
    '' #AILTON - PONTO DE ENTRADA - ( CORREÇÃO DE FLUXO )
    '' #DUVIDA - QUAL O OBJETIVO ?
    '' #ENTENDIMENTO - CARREGAR DADOS DO ARQUIVO
    x = LerXML(DocXML)
    
    
End If

AbrirConexao

txtIdentifCNPJ = False

''' #AILTON - PONTO DE ENTRADA - ( CORREÇÃO DE FLUXO )
'x = LerXML(DocXML)

'MsgBox XMLNumNF & " " & XMLCNPJEmi & " " & XMLValNF & " " & XMLBCICMS & " " & XMLPerICMS & " " & XMLValICMS & " " & Chr(10) & XMLDTEmi & " " & XMLChave


'' #AILTON - qryUpdateFornecedoresValidos
'' #ENTENDIMENTO - BLOQUEIO POR FORNECEDOR - ( VAMOS TRABALHAR APENAS COM EMITENTE )
strSql = "SELECT Clientes.CNPJ_CPF, Clientes.CÓDIGOClientes AS ID_Cad FROM Clientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "');"


'SaveSQLString strSQL
rsCad.CursorType = adOpenKeyset
rsCad.LockType = adLockOptimistic
rsCad.CursorLocation = adUseClient
rsCad.Open strSql, CNN, , , adCmdText
If rsCad.RecordCount = 0 Then
    MsgBox "Fornecedor não cadastrado !" & Chr(10) & "CNPJ: " & XMLCNPJEmi, vbInformation, "Atenção"
    rsCad.Close
    Exit Sub
Else
    '' #AILTON - ID_FORNECEDOR - qryUpdateFornecedoresValidos
    IDCadFor = rsCad!ID_Cad
End If

'' #AILTON - qryExpurgoDeComprasJaCadastradas
'' #DUVIDA - QUAL O OBJETIVO ?
'' #ENTENDIMENTO - NÃO GERAR DUPLICIDADE NO CADASTRO DE COMPRA VERIFICANDO A EXISTENCIA DA MESMA PELOS SEGUINTES CAMPOS: XMLCNPJEmi E XMLNumNF
strSql = "SELECT Clientes.CNPJ_CPF, tblCompraNF.ID_Forn_CompraNF, tblCompraNF.NumNF_CompraNF, tblCompraNF.DTEntd_CompraNF FROM tblCompraNF INNER JOIN Clientes ON tblCompraNF.ID_Forn_CompraNF = Clientes.CÓDIGOClientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "') AND (tblCompraNF.NumNF_CompraNF=" & XMLNumNF & ");"

'SaveSQLString strSQL

rsCompraNF.CursorType = adOpenKeyset
rsCompraNF.LockType = adLockOptimistic
rsCompraNF.CursorLocation = adUseClient
rsCompraNF.Open strSql, CNN, , , adCmdText
If rsCompraNF.RecordCount > 0 Then
    'MsgBox Me.Finalidade.Column(1) & " já importad" & IIf(Me.Finalidade.Column(0) = 0, "o", "a") & "  !", vbInformation, "Atenção"
    strXMLjaImp = strXMLjaImp & XMLNumNF & " - Data Entrada: " & rsCompraNF!DTEntd_CompraNF & Chr(10)
    'Contador XML's já importados
    CountXMLImp = CountXMLImp + 1
    rsCompraNF.Close
    Exit Sub
End If
rsCompraNF.Close



'Contador XML's novos
CountXMLNovo = CountXMLNovo + 1

'' #AILTON
'' #tblTipos - 0;CT-e
'' Finalidade - 0;CT-e;1;NF-e Importação;2;NF-e Consumo;3;NF-e com código Sisparts;4;NF-e Retorno Armazém;5;NF-e;6;NF-e Transferência com código Sisparts;7;NF-e Transferência Uso/Consumo com código Sisparts
If Me.Finalidade = 0 Then 'CT-e

       
    '' #Ailton - qryUpdateCFOP_PSP
    '' #ENTENDIMENTO_01 - ATUALIZAÇÃO DE "ID_NatOp_CompraNF" E "CFOP_CompraNF" COM BASE "ID_EMPRESA"
    If Me.FiltroFil = "PSP" Then
        
        '' #DUVIDA - QUAL O OBJETIVO DESSE CRITERIO ?
        '' #ENTENDIMENTO_02 - SE ACHARMOS O "txtCNPJRem" A NATUREZA DA OPERAÇÃO É 1.353 SE NÃO SERÁ 2.353
        If Me.txtUFEmit = DLookup("[UF_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'") Then
            
            '' cfop - local
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        ElseIf Me.txtUFEmit <> DLookup("[UF_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'") Then
            
            '' cfop - fora do estado
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        End If
        
    '' #Ailton - qryUpdateCFOP_PSC_PES
    '' #ENTENDIMENTO_01 - ATUALIZAÇÃO DE "ID_NatOp_CompraNF" E "CFOP_CompraNF" COM BASE "ID_EMPRESA"
    '' #ENTENDIMENTO_02 - SE O XMLCFOP FOR 5353 NATUREZA DA OPERAÇÃO PARA PESQUISA SERÁ 1.353 SE NÃO SERÁ 2.353 COM O "FiltroFil" QUE VEM DO CADASTRO
    
    ElseIf Me.FiltroFil = "PSC" Then
    
         If XMLCFOP = "5353" Then '1353
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        ElseIf XMLCFOP = "6353" Then '2353
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        End If
    ElseIf Me.FiltroFil = "PES" Then
    
         If XMLCFOP = "5353" Then '1353
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        ElseIf XMLCFOP = "6353" Then '2353
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        End If
    End If
    

    
    rsCompraNF.CursorType = adOpenKeyset
    rsCompraNF.LockType = adLockOptimistic
    rsCompraNF.CursorLocation = adUseClient
    
    '' #DUVIDA - QUAL O OBJETIVO ?
    rsCompraNF.Open "SELECT TOP 1 * FROM tblCompraNF", CNN, , , adCmdText

    rsCompraNF.AddNew '' CONHECIMENTO
    
    
    '' #AILTON - qryInsertCompra - #VALOR_VARIAVEL
    rsCompraNF!Fil_CompraNF = Me.FiltroFil
    rsCompraNF!NumNF_CompraNF = XMLNumNF
    rsCompraNF!NumPed_CompraNF = Format(IIf(IsNull(DMax("NumPed_CompraNF", "tblCompraNF")), "000001", DMax("NumPed_CompraNF", "tblCompraNF") + 1), "000000") '' CONTROLE DO NUMERO DO PEDIDO
    rsCompraNF!ID_NatOp_CompraNF = Me.FiltroCFOP.Column(0)
    rsCompraNF!CFOP_CompraNF = Me.FiltroCFOP.Column(1)
    rsCompraNF!Serie_CompraNF = XMLSerie
    rsCompraNF!DTEmi_CompraNF = XMLDTEmi
    rsCompraNF!DTEntd_CompraNF = Date
    rsCompraNF!ID_Forn_CompraNF = IDCadFor
    rsCompraNF!BaseCalcICMS_CompraNF = XMLBCICMS
    rsCompraNF!VTotICMS_CompraNF = XMLValICMS
    rsCompraNF!VTotProd_CompraNF = XMLValNF
    rsCompraNF!VTotNF_CompraNF = XMLValNF
    rsCompraNF!ChvAcesso_CompraNF = XMLChave
    
    '' #AILTON - qryInsertCompra - #VALOR_FIXO
    rsCompraNF!ModeloDoc_CompraNF = 57
    rsCompraNF!TPNF_CompraNF = 1
    rsCompraNF!Sit_CompraNF = 6
    rsCompraNF!VTotServ_CompraNF = 0
    rsCompraNF!BaseCalcICMSSubsTrib_CompraNF = 0
    rsCompraNF!VTotICMSSubsTrib_CompraNF = 0
    rsCompraNF!VTotFrete_CompraNF = 0
    rsCompraNF!VTotSeguro_CompraNF = 0
    rsCompraNF!VTotOutDesp_CompraNF = 0
    rsCompraNF!VTotIPI_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    rsCompraNF!TxDesc_CompraNF = 0
    rsCompraNF!VTotDesc_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    
    rsCompraNF.Update
    
    '' #AILTON - qryInsertCompra - ( ID_COMPRA )
    IDCompraNF = rsCompraNF!ID_CompraNF
        
    rsCompraNF.Close
    rsCad.Close
    
    rsCompraNFItem.CursorType = adOpenKeyset
    rsCompraNFItem.LockType = adLockOptimistic
    rsCompraNFItem.CursorLocation = adUseClient
    
    '' #DUVIDA - QUAL O OBJETIVO ?
    rsCompraNFItem.Open "SELECT TOP 1 * FROM tblCompraNFItem", CNN, , , adCmdText
    rsCompraNFItem.AddNew
    
    '' #AILTON - qryInsertCompraItens
    rsCompraNFItem!ID_CompraNF_CompraNFItem = IDCompraNF
    rsCompraNFItem!ID_Prod_CompraNFItem = DLookup("[Código]", "[Cadastro de Produtos]", "[Modelo]='TRANSPORTE'")
    rsCompraNFItem!Almox_CompraNFItem = Forms!frmCompraNF_ImpXML!FiltroAlmox
    rsCompraNFItem!VUnt_CompraNFItem = XMLValNF
    rsCompraNFItem!ICMS_CompraNFItem = XMLPerICMS
    rsCompraNFItem!CFOP_CompraNFItem = Me.FiltroCFOP.Column(1)
    rsCompraNFItem!ST_CompraNFItem = "0" & Me.FiltroCFOP.Column(4)
    rsCompraNFItem!DTEmi_CompraNFItem = XMLDTEmi
    rsCompraNFItem!Num_CompraNFItem = XMLNumNF
    rsCompraNFItem!VTot_CompraNFItem = XMLValNF
    rsCompraNFItem!DebICMS_CompraNFItem = XMLValICMS
    rsCompraNFItem!VTotBaseCalcICMS_CompraNFItem = XMLBCICMS
    rsCompraNFItem!ID_NatOp_CompraNFItem = Me.FiltroCFOP.Column(0)
    rsCompraNFItem!STPIS_CompraNFItem = Me.FiltroCFOP.Column(6)
    rsCompraNFItem!STCOFINS_CompraNFItem = Me.FiltroCFOP.Column(6)
    rsCompraNFItem!STIPI_CompraNFItem = Me.FiltroCFOP.Column(5)
    
    '' #AILTON - qryInsertCompraItens - #VALOR_FIXO
    rsCompraNFItem!Item_CompraNFItem = 1
    rsCompraNFItem!ID_Grade_CompraNFItem = 1
    rsCompraNFItem!QtdFat_CompraNFItem = 1
    rsCompraNFItem!ISS_CompraNFItem = 0
    rsCompraNFItem!IPI_CompraNFItem = 0
    rsCompraNFItem!FlagEst_CompraNFItem = 0
    rsCompraNFItem!BaseCalcICMS_CompraNFItem = 100
    rsCompraNFItem!IseICMS_CompraNFItem = 0
    rsCompraNFItem!OutICMS_CompraNFItem = 0
    rsCompraNFItem!BaseCalcIPI_CompraNFItem = 0
    rsCompraNFItem!DebIPI_CompraNFItem = 0
    rsCompraNFItem!IseIPI_CompraNFItem = 0
    rsCompraNFItem!OutIPI_CompraNFItem = 0
    rsCompraNFItem!TxMLSubsTrib_CompraNFItem = 0
    rsCompraNFItem!TxIntSubsTrib_CompraNFItem = 0
    rsCompraNFItem!TxExtSubsTrib_CompraNFItem = 0
    rsCompraNFItem!BaseCalcICMSSubsTrib_CompraNFItem = 0
    rsCompraNFItem!VTotICMSSubsTrib_compranfitem = 0
    rsCompraNFItem!VTotFrete_CompraNFItem = 0
    rsCompraNFItem!VTotDesc_CompraNFItem = 0
    rsCompraNFItem!VTotFrete_CompraNFItem = 0
    rsCompraNFItem!VTotPIS_CompraNFItem = 0
    rsCompraNFItem!VTotBaseCalcPIS_CompraNFItem = 0
    rsCompraNFItem!VTotCOFINS_CompraNFItem = 0
    rsCompraNFItem!VTotBaseCalcCOFINS_CompraNFItem = 0
    rsCompraNFItem!VTotIseICMS_CompraNFItem = 0
    rsCompraNFItem!VTotBaseCalcCOFINS_CompraNFItem = 0
    rsCompraNFItem!SNCredICMS_CompraNFItem = 0
    rsCompraNFItem!VTotSNCredICMS_CompraNFItem = 0
    rsCompraNFItem!VTotSeg_CompraNFItem = 0
    rsCompraNFItem!VTotOutDesp_CompraNFItem = 0
    
    
    rsCompraNFItem.Update
    rsCompraNFItem.Close
    
    Me.txtCNPJRem = ""
    Me.txtUFEmit = ""
    i = 0
    txtContaUFEmit = 0
    

'' #AILTON - qryUpdateCFOP
'' #tblTipos - 1 - NF-e Importação / 3 - NF-e com código Sisparts / 4 - NF-e Retorno Armazém / 5 - NF-e / 6 - NF-e Transferência com código Sisparts / 7 - NF-e Transferência Uso/Consumo com código Sisparts
ElseIf Me.Finalidade = 1 Or Me.Finalidade = 3 Or Me.Finalidade = 4 Or Me.Finalidade = 5 Or Me.Finalidade = 6 Or Me.Finalidade = 7 Then
    If Me.Finalidade = 6 Then '' #AILTON - NORMAL
        Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.152' and [Fil_NatOper]='" & Me.FiltroFil & "'")
        Me.FiltroCFOP.Requery
    ElseIf Me.Finalidade = 7 Then '' #AILTON - CONSUMO
        Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.557' and [Fil_NatOper]='" & Me.FiltroFil & "'")
        Me.FiltroCFOP.Requery
    End If
        
    
    CurrentDb.Execute "DELETE * FROM 05_XML_ICMS"
    CurrentDb.Execute "DELETE * FROM 05_XML_IPI"
    
    
    rsCompraNF.CursorType = adOpenKeyset
    rsCompraNF.LockType = adLockOptimistic
    rsCompraNF.Open "SELECT TOP 1 * FROM tblCompraNF ORDER BY NumPed_CompraNF DESC;", CNN
    
    
    Dim NumPed As Long
    NumPed = rsCompraNF!NumPed_CompraNF
    
    '' CONHECIMENTO
    rsCompraNF.AddNew
    
    '' #AILTON - qryInsertCompra - #VALOR_VARIAVEL
    rsCompraNF!Fil_CompraNF = Me.FiltroFil
    rsCompraNF!NumNF_CompraNF = XMLNumNF
    rsCompraNF!NumPed_CompraNF = Format(NumPed + 1, "000000")
    rsCompraNF!Serie_CompraNF = XMLSerie
    rsCompraNF!ID_NatOp_CompraNF = Me.FiltroCFOP.Column(0)
    rsCompraNF!CFOP_CompraNF = Me.FiltroCFOP.Column(1)
    rsCompraNF!DTEmi_CompraNF = XMLDTEmi
    rsCompraNF!DTEntd_CompraNF = XMLDTEmi
    
    '' #BUG - PONTO DE QUEBRA
    rsCompraNF!Obs_CompraNF = Mid(XMLinfAdFisco, 1, 255)
        
    Dim parts() As String
    Dim NumAtual As String
   
    '' #BUG - PONTO DE QUEBRA
    '' #PEGAR APENAS O NUMERO DO PEDIDO - #PONTO DE ATENÇÃO
    If XMLinfAdFisco <> "" Then
        parts = Split((XMLinfAdFisco), ",")
        NumAtual = parts(LBound(parts))
        
'        tNumAtual = Replace(parts(LBound(Split((XMLinfAdFisco), ","))), "Pedido:", "")
        
        If Forms!frmCompraNF_ImpXML!Finalidade = 4 Then
            rsCompraNF!IDVD_CompraNF = Replace(NumAtual, "Pedido:", "")
        End If
    End If
    'rsCompraNF!IDVD_CompraNF = XMLinfAdFisco
        
    rsCompraNF!ID_Forn_CompraNF = IDCadFor
    
    '' #AILTON - AJUSTE 1.00
    If Forms!frmCompraNF_ImpXML!Finalidade = 4 Then
        rsCompraNF!VTotProd_CompraNF = XMLvNF
        rsCompraNF!VTotNF_CompraNF = XMLvNF
        rsCompraNF!HoraEntd_CompraNF = XMLdhSaiEnt
        sCompraNF!Sit_CompraNF = 6
    Else
        rsCompraNF!VTotProd_CompraNF = 0
        rsCompraNF!VTotNF_CompraNF = 0
        rsCompraNF!Sit_CompraNF = 5
    End If
    
'    '' #AILTON - AJUSTE 1.01
'    If Forms!frmCompraNF_ImpXML!Finalidade = 4 Then
'        rsCompraNF!Sit_CompraNF = 6
'    Else
'        rsCompraNF!Sit_CompraNF = 5
'    End If
    
    rsCompraNF!ChvAcesso_CompraNF = XMLChave
    rsCompraNF!BaseCalcICMSSubsTrib_CompraNF = 0
    rsCompraNF!VTotICMSSubsTrib_CompraNF = 0
    
'    If EstáAberto("formVendas") Then
'        rsCompraNF!VTotFrete_CompraNF = Forms!formVendas!Transporte * -1
'    Else
'        rsCompraNF!VTotFrete_CompraNF = 0
'    End If
    
    
    '' #AILTON - qryInsertCompraItens - #VALOR_FIXO
    
    rsCompraNF!VTotFrete_CompraNF = 0
    
    rsCompraNF!ModeloDoc_CompraNF = 55
    rsCompraNF!TPNF_CompraNF = 1
    rsCompraNF!BaseCalcICMS_CompraNF = 0
    rsCompraNF!VTotICMS_CompraNF = 0
    rsCompraNF!VTotServ_CompraNF = 0
    rsCompraNF!VTotSeguro_CompraNF = 0
    rsCompraNF!VTotOutDesp_CompraNF = 0
    rsCompraNF!VTotIPI_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    rsCompraNF!TxDesc_CompraNF = 0
    rsCompraNF!VTotDesc_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    
    
    rsCompraNF.Update       ' Salva alterações.
    
    '' #AILTON - qryInsertCompra - ( ID_COMPRA )
    IDCompraNF = rsCompraNF!ID_CompraNF
    Me.ID = IDCompraNF
    
    If Me.Finalidade = 1 Or Me.Finalidade = 3 Or Me.Finalidade = 4 Or Me.Finalidade = 5 Or Me.Finalidade = 6 Or Me.Finalidade = 7 Then
        
        Imp_Prod_ICMS (DocXML)
        Imp_Prod_IPI (DocXML)
        
        
        '' #AILTON - qryInsertCompraItens
        PegaTagProdXML DocXML, IDCompraNF, IDCadFor
        
    End If

    rsCompraNF.Close
    rsCad.Close
 
'    If EstáAberto("formVendas") Then
'    Else
'        DoCmd.SetWarnings False
'        DoCmd.OpenQuery "qryCompraNF_ImpXML_BCICMS"
'        DoCmd.OpenQuery "qryCompraNF_ImpXML_IPI"
'        DoCmd.SetWarnings True
'    End If
  
    '' #AILTON - PENDENTE
    DoCmd.OpenQuery "qryCompraNF_ImpXML_BCICMS"
    DoCmd.OpenQuery "qryCompraNF_ImpXML_IPI"
  
  
  
'' #AILTON - qryUpdateCFOP
'' #tblTipos - 2 - NF-e Consumo
ElseIf Me.Finalidade = 2 Then
    CurrentDb.Execute "DELETE * FROM 04_XML_prod"
    CurrentDb.Execute "DELETE * FROM 05_XML_ICMS"
    CurrentDb.Execute "DELETE * FROM 05_XML_IPI"
    
    rsCompraNF.CursorType = adOpenKeyset
    rsCompraNF.LockType = adLockOptimistic
    rsCompraNF.Open "SELECT TOP 1 * FROM tblCompraNF ORDER BY NumPed_CompraNF DESC;", CNN
    NumPed = rsCompraNF!NumPed_CompraNF
    
    
    rsCompraNF.AddNew
    
    '' #AILTON - qryInsertCompra
    rsCompraNF!Fil_CompraNF = Me.FiltroFil
    rsCompraNF!NumNF_CompraNF = XMLNumNF
    rsCompraNF!NumPed_CompraNF = Format(NumPed + 1, "000000")
    rsCompraNF!Serie_CompraNF = XMLSerie
    rsCompraNF!ID_NatOp_CompraNF = Me.FiltroCFOP.Column(0)
    rsCompraNF!CFOP_CompraNF = Me.FiltroCFOP.Column(1)
    rsCompraNF!DTEmi_CompraNF = XMLDTEmi
    rsCompraNF!DTEntd_CompraNF = XMLDTEmi
    rsCompraNF!ID_Forn_CompraNF = IDCadFor
    
'    '' #AILTON - NÃO ENTENDI O MOTIVO DISSO!
'    If EstáAberto("formVendas") Then
'        rsCompraNF!BaseCalcICMS_CompraNF = 0
'    Else
'        rsCompraNF!BaseCalcICMS_CompraNF = 0
'    End If
    
    
    '' #AILTON - qryInsertCompra - #VALOR_FIXO
    rsCompraNF!BaseCalcICMS_CompraNF = 0    '' <<<<<< ESTOU USANDO AQUI
    rsCompraNF!ModeloDoc_CompraNF = 55
    rsCompraNF!TPNF_CompraNF = 1
    rsCompraNF!VTotICMS_CompraNF = 0
    rsCompraNF!VTotServ_CompraNF = 0
    rsCompraNF!VTotProd_CompraNF = 0
    rsCompraNF!VTotNF_CompraNF = 0
    rsCompraNF!Sit_CompraNF = 5
    rsCompraNF!ChvAcesso_CompraNF = XMLChave
    rsCompraNF!BaseCalcICMSSubsTrib_CompraNF = 0
    rsCompraNF!VTotICMSSubsTrib_CompraNF = 0
    rsCompraNF!VTotFrete_CompraNF = 0
    rsCompraNF!VTotSeguro_CompraNF = 0
    rsCompraNF!VTotOutDesp_CompraNF = 0
    rsCompraNF!VTotIPI_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    rsCompraNF!TxDesc_CompraNF = 0
    rsCompraNF!VTotDesc_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
    rsCompraNF!VTotISS_CompraNF = 0
        
    
    rsCompraNF.Update       ' Salva alterações.
    
    '' #AILTON - qryInsertCompra - ( ID_COMPRA )
    IDCompraNF = rsCompraNF!ID_CompraNF
    Me.ID = IDCompraNF
    
'    If Me.Finalidade = 2 Then
'        Imp_Prod_NF (DocXML)
'        Imp_Prod_ICMS (DocXML)
'        Imp_Prod_IPI (DocXML)
'
'        '' #Ailton - qryInsertProdutoConsumo
'        CadProdXMLConsumo (IDCadFor)
'
'        '' #Ailton - qryInsertCompraItens
'        PegaTagProdXML DocXML, IDCompraNF, IDCadFor
'
'    End If
       
End If


strXMLnovo = strXMLnovo & XMLNumNF & Chr(10)

'MsgBox "XML Importado com Sucesso !!!", vbExclamation + vbOKOnly, "Atenção"
verificaImpXML = True

TrataErro_Exit:
    Exit Sub

TrataErro:
    MsgBox err.description
    DoCmd.Hourglass False 'Mouse Ampulheta
    Resume Next 'TrataErro_Exit
End Sub


