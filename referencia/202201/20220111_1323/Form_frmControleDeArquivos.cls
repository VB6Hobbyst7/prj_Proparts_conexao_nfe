VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmControleDeArquivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'' #20220110_qryComprasItem_Update_AjustesCampos
Private Const qryComprasItem_Update_AjustesCampos As String = _
            "UPDATE tblCompraNF " & _
            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
            "SET tblCompraNFItem.ID_Grade_CompraNFItem = '1' " & _
            "   ,tblCompraNFItem.FlagEst_CompraNFItem = '1' " & _
            "   ,tblCompraNFItem.BaseCalcICMS_CompraNFItem = '100' " & _
            "   ,tblCompraNFItem.BaseCalcICMSSubsTrib_CompraNFItem = '0' " & _
            "   ,tblCompraNFItem.ID_NatOp_CompraNFItem = [tblCompraNF].[ID_NatOp_CompraNF] " & _
            "   ,tblCompraNFItem.CFOP_CompraNFItem = [tblCompraNF].[CFOP_CompraNF];"


'' #20220110_qryComprasItem_Update_STs
Private Const qryComprasItem_Update_STs As String = _
            "UPDATE tblCompraNFItem " & _
            "SET tblCompraNFItem.ST_CompraNFItem = ""1"" & DLookUp(""[STICMS_NatOper]"", ""[tmpNatOp]"", ""[ID_NatOper] = "" & [tblCompraNFItem].[ID_NatOp_CompraNFItem] & """") " & _
            "   ,tblCompraNFItem.STCOFINS_CompraNFItem = DLookUp(""[STPC_NatOper]"", ""[tmpNatOp]"", ""[ID_NatOper] = "" & [tblCompraNFItem].[ID_NatOp_CompraNFItem] & """") " & _
            "   ,tblCompraNFItem.STPIS_CompraNFItem = DLookUp(""[STPC_NatOper]"", ""[tmpNatOp]"", ""[ID_NatOper] = "" & [tblCompraNFItem].[ID_NatOp_CompraNFItem] & """") " & _
            "   ,tblCompraNFItem.STIPI_CompraNFItem = DLookUp(""[STIPI_NatOper]"", ""[tmpNatOp]"", ""[ID_NatOper] = "" & [tblCompraNFItem].[ID_NatOp_CompraNFItem] & """");"

'' #20220111_update_Almox_CompraNFItem
Private Const qryComprasItem_Update_Almox_CompraNFItem_55_1907 As String = _
            "UPDATE tblCompraNF " & _
            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
            "SET tblCompraNFItem.Almox_CompraNFItem = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Almox_CompraNFItem|55|1.907'"") " & _
            "WHERE (((tblCompraNF.CFOP_CompraNF) = ""1.907"") AND ((tblCompraNF.ModeloDoc_CompraNF) = ""55""));"

'' #20220111_update_Almox_CompraNFItem
Private Const qryComprasItem_Update_Almox_CompraNFItem_55_2152_PSC As String = _
            "UPDATE tblCompraNF " & _
            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
            "SET tblCompraNFItem.Almox_CompraNFItem = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Almox_CompraNFItem|55|2.152|PSC'"") " & _
            "WHERE (((tblCompraNF.CFOP_CompraNF) = ""2.152"") AND ((tblCompraNF.Fil_CompraNF) = ""PSC"")    AND ((tblCompraNF.ModeloDoc_CompraNF) = ""55"") );"

'' #20220111_update_Almox_CompraNFItem
Private Const qryComprasItem_Update_Almox_CompraNFItem_55_2152_PSP As String = _
            "UPDATE tblCompraNF " & _
            "INNER JOIN tblCompraNFItem ON tblCompraNF.ChvAcesso_CompraNF = tblCompraNFItem.ChvAcesso_CompraNF " & _
            "SET tblCompraNFItem.Almox_CompraNFItem = DLookUp(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro] = 'Almox_CompraNFItem|55|2.152|PSP'"") " & _
            "WHERE (((tblCompraNF.CFOP_CompraNF) = ""2.152"") AND ((tblCompraNF.Fil_CompraNF) = ""PSP"")    AND ((tblCompraNF.ModeloDoc_CompraNF) = ""55"") );"




Private Sub cmdDadosGerais_Click()
Dim pArquivos As Collection: Set pArquivos = New Collection
Dim strTitulo As String: strTitulo = "Processar Dados Gerais"
Dim retVal As Variant: retVal = MsgBox("Deseja iniciar o processamento de arquivo para dados gerais?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        processarDadosGerais pArquivos
        Me.Requery
        MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
    End If

Debug.Print strTitulo & " - Concluido!"
End Sub

Private Sub cmdArquivosPendentes_Click()
Dim strTitulo As String: strTitulo = "Processar arquivos pendentes"
Dim retVal As Variant: retVal = MsgBox("Deseja iniciar o processamento dos arquivos pendentes?", vbQuestion + vbOKCancel, strTitulo)

Dim qryProcessos() As Variant: qryProcessos = Array( _
                                                    qryComprasItem_Update_AjustesCampos, _
                                                    qryComprasItem_Update_STs _
                                                    )
    If retVal = vbOK Then
        If (DCount("[id]", "tblDadosConexaoNFeCTe", "[registroProcessado]=0") > 0) Then
            processarArquivosPendentes
            executarComandos qryProcessos
            Me.Requery
            MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
        Else
            MsgBox "Não há dados para processamentos.", vbCritical + vbOKOnly, strTitulo
        End If
    End If

Debug.Print strTitulo & " - Concluido!"
End Sub


Private Sub cmdEnviarDadosServidor_Click()
Dim strTitulo As String: strTitulo = "Enviar dados para o servidor"
Dim retVal As Variant: retVal = MsgBox("Deseja iniciar o envio de dados para o servidor?", vbQuestion + vbOKCancel, strTitulo)
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColetaAcoes'")

    If retVal = vbOK Then
        If (DCount("[id]", "tblDadosConexaoNFeCTe", "[registroProcessado]=1") > 0) Then
            CadastroDeComprasEmServidor
            MoverArquivosProcessados
            
            '' LANÇAMENTO
            gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
            
            '' MANIFESTO
            gerarArquivosJson opManifesto, , strCaminhoAcoes
            
            Me.Requery
            MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
        Else
            MsgBox "Não há dados para processamentos.", vbCritical + vbOKOnly, strTitulo
        End If
    End If

Debug.Print strTitulo & " - Concluido!"
End Sub

Private Sub cmdLancamento_Click()
Dim strTitulo As String: strTitulo = "Gerar Json de Lancamento."
Dim retVal As Variant: retVal = MsgBox("Deseja gerar Json de desse item.?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        If Me.txtRegistroValido.value = 1 Then
            If Me.txtChvAcesso.value <> "" Then gerarArquivosJson opFlagLancadaERP, Me.txtChvAcesso.value, Me.txtCaminhoPadrao.value
            MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
            
            Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
            
            Me.Requery
            MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
        Else
            MsgBox "Disponivel apenas para ""Registros valido - OK"" ", vbInformation + vbOKOnly, strTitulo
        End If
    End If

Debug.Print strTitulo & " - Concluido!"
End Sub

Private Sub cmdLancamentoTodos_Click()
Dim strTitulo As String: strTitulo = "Gerar Json de todos os lancamentos."
Dim retVal As Variant: retVal = MsgBox("Deseja gerar Json de todos os lancamentos.?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        gerarArquivosJson opFlagLancadaERP, , Me.txtCaminhoPadrao.value
        Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
        MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
    End If

Debug.Print strTitulo & " - Concluido!"
End Sub

Private Sub cmdManifesto_Click()
Dim strTitulo As String: strTitulo = "Gerar Json de Manifesto."
Dim retVal As Variant: retVal = MsgBox("Deseja gerar Json de desse item.?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        If Me.txtRegistroValido.value = 1 Then
            If Me.txtChvAcesso.value <> "" Then
                gerarArquivosJson opManifesto, Me.txtChvAcesso.value, Me.txtCaminhoPadrao.value
                Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
                Me.Requery
                MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
            End If
        Else
            MsgBox "Disponivel apenas para ""Registros valido - OK"" ", vbInformation + vbOKOnly, strTitulo
        End If
    End If
    
Debug.Print strTitulo & " - Concluido!"
End Sub

Private Sub cmdManifestoTodos_Click()
Dim strTitulo As String: strTitulo = "Gerar Json de todos os manifestos."
Dim retVal As Variant: retVal = MsgBox("Deseja gerar Json de todos os manifestos.?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        gerarArquivosJson opManifesto, , Me.txtCaminhoPadrao.value
        Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
        MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
    End If
    
Debug.Print strTitulo & " - Concluido!"
End Sub

