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
        strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
    ElseIf Me.FiltroFil = "PES" Then
        strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
    Else
        strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
    End If
    'strTemp = AbreSelArquivo(hWnd, , "P:\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL\FILIAL\NF RECEBIDAS", "Todos os arquivos" & Chr(0) & "*.xml")
Else
    strTemp = AbreSelArquivo(hWnd, , "\\192.168.10.200\ProParts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL", "Todos os arquivos" & Chr(0) & "*.xml")
End If

If Trim(" " & strTemp) <> "" Then
 Me.caminho = strTemp
End If
strTemp = ""
TrataErro_Exit:
    Exit Sub

TrataErro:
    MsgBox Err.Description
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
            .InitialFileName = "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL"
        ElseIf Me.FiltroFil = "PES" Then
            .InitialFileName = "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL"
        Else
            .InitialFileName = "\\192.168.10.200\proparts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL"
        End If
        '.InitialFileName = "P:\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL\FILIAL\NF RECEBIDAS"
    Else
        .InitialFileName = "\\192.168.10.200\ProParts\Administrativo\Depto Fiscal\ARQUIVOS ELETR�NICOS - FISCAL"
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
    If Est�Aberto("formVendas") Then
        DoCmd.Close acForm, "formVendas"
    End If
End If

    CountXMLImp = 0
    CountXMLNovo = 0

If IsNull(Me.caminho) Then
    MsgBox "Selecione uma pasta ou arquivo", vbCritical, "Selecione"
    Exit Sub
End If

'Verificar se est� importando um arquivo ou uma pasta
'Vari�vel tipImpor altera de acordo com o bot�o clicado, arquivo ou pasta
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
        
        'Verifica se determinado XML j� foi importado anteriormente
        'atrav�s da vari�vel CountXMLImp, que conta o n�meros de XML que j� estavam importados, na Imprimir_Click
        If CountXMLImp <= 0 Then
        
        'Vari�vel booleana que verifica se a rotina de importa��o foi bem sucedida
        '� definida como False no in�cio da Imprimir_Click e como True no final
            If verificaImpXML = False Then Exit Sub
            
            End If
        End If
        MyName = Dir
    Loop
        
        
End If
        'Caso o XML seja novo, mostrar� o msgbox
        
        If CountXMLNovo > 0 Then
        MsgBox "Importa��o Conclu�da!" & Chr(10) & Chr(10) & "XML's importados(N� da Nota):" & Chr(10) & Chr(10) & strXMLnovo, vbInformation, "Importa��o realizada!"
        End If
        
        'Caso j� tenha sido importado determinado XML, mostrar� o msgbox
        
        If CountXMLImp > 0 Then
        MsgBox "Os XML's abaixo j� foram importados anteriormente:(N� da Nota)" & Chr(10) & Chr(10) & strXMLjaImp, vbCritical, "Importa��o j� existente!"
        End If
        
        strXMLnovo = ""
        strXMLjaImp = ""
        DocXML = ""

Exit Sub
ErroPasta:
    MsgBox "Erro n�mero: " & Err.Number & Chr(10) & "Descri��o: " & Err.Description, vbCritical, "Erro"
    strXMLnovo = ""
    strXMLjaImp = ""
    CountXMLImp = 0
    CountXMLNovo = 0
    DocXML = ""
End Sub

Private Sub Imprimir_Click() '' #AILTON - COMPRANF
On Error GoTo TrataErro
Dim strSQL As String

Dim rsCad As ADODB.Recordset: Set rsCad = New ADODB.Recordset
Dim rsCompraNF As ADODB.Recordset: Set rsCompraNF = New ADODB.Recordset
Dim rsCompraNFItem As ADODB.Recordset: Set rsCompraNFItem = New ADODB.Recordset

Dim IDCompraNF As Double
Dim IDCadFor As Long

verificaImpXML = False

If IsNull(Me.Finalidade) Or IsNull(Me.caminho) Then
    MsgBox "Digite finalidade e selecione arquivo!", vbCritical, "Aten��o"
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
        MsgBox "Digite CFOP!", vbCritical, "Aten��o"
        Me.FiltroCFOP.SetFocus
        Exit Sub
    End If
    
    '' #AILTON
    '' qrySelectEmpresa_FiltroFil
    If IsNull(Me.FiltroFil) Then
        MsgBox "Selecione a Empresa!", vbCritical, "Aten��o"
        Me.FiltroFil.SetFocus
        Exit Sub
    End If
    
    If IsNull(Me.FiltroAlmox) Then
        MsgBox "Selecione almoxarifado!", vbCritical, "Aten��o"
        Me.FiltroAlmox.SetFocus
        Exit Sub
    End If
    
    '' #AILTON - PONTO DE ENTRADA - ( CORRE��O DE FLUXO )
    '' #DUVIDA - QUAL O OBJETIVO ?
    '' #ENTENDIMENTO - CARREGAR DADOS DO ARQUIVO
    x = LerXML(DocXML)
    
    
End If

AbrirConexao

txtIdentifCNPJ = False

''' #AILTON - PONTO DE ENTRADA - ( CORRE��O DE FLUXO )
'x = LerXML(DocXML)

'MsgBox XMLNumNF & " " & XMLCNPJEmi & " " & XMLValNF & " " & XMLBCICMS & " " & XMLPerICMS & " " & XMLValICMS & " " & Chr(10) & XMLDTEmi & " " & XMLChave


'' #AILTON - qryDadosGerais_Update_FornecedoresValidos
'' ID_Forn_CompraNF
'' #ENTENDIMENTO - BLOQUEIO POR FORNECEDOR - ( VAMOS TRABALHAR APENAS COM EMITENTE )
strSQL = "SELECT Clientes.CNPJ_CPF, Clientes.C�DIGOClientes AS ID_Cad FROM Clientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "');"


'' IDCadFor
'SaveSQLString strSQL
rsCad.CursorType = adOpenKeyset
rsCad.LockType = adLockOptimistic
rsCad.CursorLocation = adUseClient
rsCad.Open strSQL, CNN, , , adCmdText
If rsCad.RecordCount = 0 Then
    MsgBox "Fornecedor n�o cadastrado !" & Chr(10) & "CNPJ: " & XMLCNPJEmi, vbInformation, "Aten��o"
    rsCad.Close
    Exit Sub
Else
    '' #AILTON - ID_FORNECEDOR - qryDadosGerais_Update_FornecedoresValidos
    IDCadFor = rsCad!ID_Cad
End If

'' #AILTON - qryExpurgoDeComprasJaCadastradas
'' #DUVIDA - QUAL O OBJETIVO ?
'' #ENTENDIMENTO - N�O GERAR DUPLICIDADE NO CADASTRO DE COMPRA VERIFICANDO A EXISTENCIA DA MESMA PELOS SEGUINTES CAMPOS: XMLCNPJEmi E XMLNumNF
strSQL = "SELECT Clientes.CNPJ_CPF, tblCompraNF.ID_Forn_CompraNF, tblCompraNF.NumNF_CompraNF, tblCompraNF.DTEntd_CompraNF FROM tblCompraNF INNER JOIN Clientes ON tblCompraNF.ID_Forn_CompraNF = Clientes.C�DIGOClientes WHERE  (Clientes.CNPJ_CPF='" & Format(XMLCNPJEmi, "00\.000\.000/0000\-00") & "') AND (tblCompraNF.NumNF_CompraNF=" & XMLNumNF & ");"

'SaveSQLString strSQL

rsCompraNF.CursorType = adOpenKeyset
rsCompraNF.LockType = adLockOptimistic
rsCompraNF.CursorLocation = adUseClient
rsCompraNF.Open strSQL, CNN, , , adCmdText
If rsCompraNF.RecordCount > 0 Then
    'MsgBox Me.Finalidade.Column(1) & " j� importad" & IIf(Me.Finalidade.Column(0) = 0, "o", "a") & "  !", vbInformation, "Aten��o"
    strXMLjaImp = strXMLjaImp & XMLNumNF & " - Data Entrada: " & rsCompraNF!DTEntd_CompraNF & Chr(10)
    'Contador XML's j� importados
    CountXMLImp = CountXMLImp + 1
    rsCompraNF.Close
    Exit Sub
End If
rsCompraNF.Close



'Contador XML's novos
CountXMLNovo = CountXMLNovo + 1

'' #AILTON
'' #tblTipos - 0;CT-e
'' Finalidade - 0;CT-e;1;NF-e Importa��o;2;NF-e Consumo;3;NF-e com c�digo Sisparts;4;NF-e Retorno Armaz�m;5;NF-e;6;NF-e Transfer�ncia com c�digo Sisparts;7;NF-e Transfer�ncia Uso/Consumo com c�digo Sisparts
If Me.Finalidade = 0 Then 'CT-e

       
    '' #Ailton - qryUpdateCFOP_PSP
    '' #ENTENDIMENTO_01 - ATUALIZA��O DE "ID_NatOp_CompraNF" E "CFOP_CompraNF" COM BASE "ID_EMPRESA"
    
    '' tblDadosConexaoNFeCTe.ID_EMPRESA
    If Me.FiltroFil = "PSP" Then
        
        '' #DUVIDA - QUAL O OBJETIVO DESSE CRITERIO ?
        '' #ENTENDIMENTO_02 - SE ACHARMOS O "txtCNPJRem" A NATUREZA DA OPERA��O � 1.353 SE N�O SER� 2.353
        If Me.txtUFEmit = DLookup("[UF_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'") Then
            
            '' cfop - local
            '' tblDadosConexaoNFeCTe.FiltroCFOP
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        ElseIf Me.txtUFEmit <> DLookup("[UF_Empresa]", "tblEmpresa", "[CNPJ_Empresa]='" & Forms!frmCompraNF_ImpXML!txtCNPJRem & "'") Then
            
            '' cfop - fora do estado
            '' tblDadosConexaoNFeCTe.FiltroCFOP
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        End If
        
    '' #Ailton - qryDadosGerais_Update_ID_NatOp_CompraNF__FiltroCFOP
    '' #ENTENDIMENTO_01 - ATUALIZA��O DE "ID_NatOp_CompraNF" E "CFOP_CompraNF" COM BASE "ID_EMPRESA"
    '' #ENTENDIMENTO_02 - SE O XMLCFOP FOR 5353 NATUREZA DA OPERA��O PARA PESQUISA SER� 1.353 SE N�O SER� 2.353 COM O "FiltroFil" QUE VEM DO CADASTRO
    
    ' PSC | 5353 | 1.353
    ' PSC | 6353 | 2.353
    
    ElseIf Me.FiltroFil = "PSC" Then
    
         If XMLCFOP = "5353" Then '1353
            
            '' tblDadosConexaoNFeCTe.FiltroCFOP
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        ElseIf XMLCFOP = "6353" Then '2353
            
            '' tblDadosConexaoNFeCTe.FiltroCFOP
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='2.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        End If
    ElseIf Me.FiltroFil = "PES" Then
    
         If XMLCFOP = "5353" Then '1353
            
            '' tblDadosConexaoNFeCTe.FiltroCFOP
            Me.FiltroCFOP = DLookup("[ID_NatOper]", "tblNatOp", "[CFOP_NatOper]='1.353' and [Fil_NatOper]='" & Me.FiltroFil & "'")
            
            Me.FiltroCFOP.Requery
        ElseIf XMLCFOP = "6353" Then '2353
            
            '' tblDadosConexaoNFeCTe.FiltroCFOP
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
    
    '' #AILTON - NumPed_CompraNF
    '' #EnviarDadosParaServidor
    rsCompraNF!NumPed_CompraNF = Format(IIf(IsNull(DMax("NumPed_CompraNF", "tblCompraNF")), "000001", DMax("NumPed_CompraNF", "tblCompraNF") + 1), "000000") '' CONTROLE DO NUMERO DO PEDIDO
    
    rsCompraNF!ID_NatOp_CompraNF = Me.FiltroCFOP.Column(0)
    rsCompraNF!CFOP_CompraNF = Me.FiltroCFOP.Column(1)
    rsCompraNF!Serie_CompraNF = XMLSerie
    rsCompraNF!DTEmi_CompraNF = XMLDTEmi
    
    '' #20210823_qryCompras_Update_Dados_NFe
    rsCompraNF!DTEntd_CompraNF = Date
    
    rsCompraNF!ID_Forn_CompraNF = IDCadFor
    rsCompraNF!BaseCalcICMS_CompraNF = XMLBCICMS
    rsCompraNF!VTotICMS_CompraNF = XMLValICMS
    
    '' #20210823_VTotProd_CompraNF
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
    
    '' #AILTON - qryComprasItens_Insert_Dados_CTeItens
    rsCompraNFItem!ID_CompraNF_CompraNFItem = IDCompraNF
    
    '' #EnviarDadosParaServidor
    '' #20210823_ID_Prod_CompraNFItem
    rsCompraNFItem!ID_Prod_CompraNFItem = DLookup("[C�digo]", "[Cadastro de Produtos]", "[Modelo]='TRANSPORTE'")
    
    '' #AILTON - qryComprasItens_Insert_Dados_CTeItens - #VALOR_VARIAVEL
    rsCompraNFItem!Almox_CompraNFItem = Forms!frmCompraNF_ImpXML!FiltroAlmox
    
    '' #20210823_VTotProd_CompraNF
    rsCompraNFItem!VUnt_CompraNFItem = XMLValNF                                 '' | OK
    rsCompraNFItem!VTot_CompraNFItem = XMLValNF                                 '' | OK
    
    rsCompraNFItem!ICMS_CompraNFItem = XMLPerICMS                               '' | carregarItens
    rsCompraNFItem!CFOP_CompraNFItem = Me.FiltroCFOP.Column(1)                  '' | tblNatOp.CFOP_NatOper
    rsCompraNFItem!ST_CompraNFItem = "0" & Me.FiltroCFOP.Column(4)              '' | tblNatOp.STICMS_NatOper
    rsCompraNFItem!DTEmi_CompraNFItem = XMLDTEmi                                '' | OK
    rsCompraNFItem!Num_CompraNFItem = XMLNumNF                                  '' | OK
    
    rsCompraNFItem!DebICMS_CompraNFItem = XMLValICMS                            '' | OK
    rsCompraNFItem!VTotBaseCalcICMS_CompraNFItem = XMLBCICMS                    '' | OK
    rsCompraNFItem!ID_NatOp_CompraNFItem = Me.FiltroCFOP.Column(0)              '' | OK
    rsCompraNFItem!STPIS_CompraNFItem = Me.FiltroCFOP.Column(6)                 '' | tblNatOp.STPC_NatOper
    rsCompraNFItem!STCOFINS_CompraNFItem = Me.FiltroCFOP.Column(6)              '' | tblNatOp.STPC_NatOper
    rsCompraNFItem!STIPI_CompraNFItem = Me.FiltroCFOP.Column(5)                 '' | tblNatOp.STIPI_NatOper
    
    '' #AILTON - qryComprasItens_Insert_Dados_CTeItens - #VALOR_FIXO
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
    rsCompraNFItem!VTotICMSSubsTrib_CompraNFItem = 0
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
'' #tblTipos - 1 - NF-e Importa��o / 3 - NF-e com c�digo Sisparts / 4 - NF-e Retorno Armaz�m / 5 - NF-e / 6 - NF-e Transfer�ncia com c�digo Sisparts / 7 - NF-e Transfer�ncia Uso/Consumo com c�digo Sisparts
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
    
    '' #EnviarDadosParaServidor
    rsCompraNF!NumPed_CompraNF = Format(NumPed + 1, "000000")
    
    rsCompraNF!Serie_CompraNF = XMLSerie
    rsCompraNF!ID_NatOp_CompraNF = Me.FiltroCFOP.Column(0)
    rsCompraNF!CFOP_CompraNF = Me.FiltroCFOP.Column(1)
    rsCompraNF!DTEmi_CompraNF = XMLDTEmi
    
    '' #20210823_qryCompras_Update_Dados_NFe
    rsCompraNF!DTEntd_CompraNF = XMLDTEmi
    
    '' #BUG - PONTO DE QUEBRA
    
    '' #EnviarDadosParaServidor
    rsCompraNF!Obs_CompraNF = Mid(XMLinfAdFisco, 1, 255)
        
    Dim parts() As String
    Dim NumAtual As String
   
    '' #BUG - PONTO DE QUEBRA
    '' #PEGAR APENAS O NUMERO DO PEDIDO - #PONTO DE ATEN��O
    
    '' #EnviarDadosParaServidor
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
    
'    If Est�Aberto("formVendas") Then
'        rsCompraNF!VTotFrete_CompraNF = Forms!formVendas!Transporte * -1
'    Else
'        rsCompraNF!VTotFrete_CompraNF = 0
'    End If
    
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
    
    
    rsCompraNF.Update       ' Salva altera��es.
    
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
 
'    If Est�Aberto("formVendas") Then
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
    
    '' #20210823_qryCompras_Update_Dados_NFe
    rsCompraNF!DTEntd_CompraNF = XMLDTEmi
    rsCompraNF!ID_Forn_CompraNF = IDCadFor
    
'    '' #AILTON - N�O ENTENDI O MOTIVO DISSO!
'    If Est�Aberto("formVendas") Then
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
        
    
    rsCompraNF.Update       ' Salva altera��es.
    
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

'MsgBox "XML Importado com Sucesso !!!", vbExclamation + vbOKOnly, "Aten��o"
verificaImpXML = True

TrataErro_Exit:
    Exit Sub

TrataErro:
    MsgBox Err.Description
    DoCmd.Hourglass False 'Mouse Ampulheta
    Resume Next 'TrataErro_Exit
End Sub


