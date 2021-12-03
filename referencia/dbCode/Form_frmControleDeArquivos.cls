<<<<<<< HEAD
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

    If retVal = vbOK Then
        If (DCount("[id]", "tblDadosConexaoNFeCTe", "[registroProcessado]=0") > 0) Then
            processarArquivosPendentes
            Me.Requery
            MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
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

=======
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

Private Sub cmdDadosGerais_Click()
Dim strTitulo As String: strTitulo = "Processar Dados Gerais"
Dim retVal As Variant: retVal = MsgBox("Deseja iniciar o processamento de arquivo para dados gerais ?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        processarDadosGerais
        Me.Requery
        MsgBox "Concluido!", vbOKOnly, strTitulo
    End If

End Sub

Private Sub cmdArquivosPendentes_Click()
Dim strTitulo As String: strTitulo = "Processar arquivos pendentes"
Dim retVal As Variant: retVal = MsgBox("Deseja iniciar o processamento dos arquivos pendentes ?", vbQuestion + vbOKCancel, strTitulo)

    If retVal = vbOK Then
        processarArquivosPendentes
        Me.Requery
        MsgBox "Concluido!", vbOKOnly, strTitulo
    End If

End Sub


Private Sub cmdEnviarDadosServidor_Click()
Dim strTitulo As String: strTitulo = "Enviar dados para o servidor"
Dim retVal As Variant: retVal = MsgBox("Deseja iniciar o envio de dados para o servidor ?", vbQuestion + vbOKCancel, strTitulo)
Dim strCaminhoAcoes As String: strCaminhoAcoes = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='caminhoDeColeta'")


    If retVal = vbOK Then
        CadastroDeComprasEmServidor
        tratamentoDeArquivosValidos
        tratamentoDeArquivosInvalidos
        
        '' LANÇAMENTO
        gerarArquivosJson opFlagLancadaERP, , strCaminhoAcoes
        
        '' MANIFESTO
        gerarArquivosJson opManifesto, , strCaminhoAcoes
        
        Me.Requery
        MsgBox "Concluido!", vbOKOnly, strTitulo

    End If

End Sub

Private Sub cmdLancamento_Click()
Dim strTitulo As String: strTitulo = "Gerar Json de Lancamento."

    If Me.txtRegistroValido.value = 1 Then
        If Me.txtChvAcesso.value <> "" Then gerarArquivosJson opFlagLancadaERP, Me.txtChvAcesso.value, Me.txtCaminhoPadrao.value
        MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
        
        Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
        
        Me.Requery
        MsgBox "Concluido!", vbOKOnly, strTitulo
    Else
        MsgBox "Disponivel apenas para ""Registros valido - OK"" ", vbInformation + vbOKOnly, strTitulo
    End If

End Sub

Private Sub cmdLancamentoTodos_Click()

    gerarArquivosJson opFlagLancadaERP, , Me.txtCaminhoPadrao.value
    Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus

End Sub

Private Sub cmdManifesto_Click()
Dim strTitulo As String: strTitulo = "Gerar Json de Manifesto."

    If Me.txtRegistroValido.value = 1 Then
        If Me.txtChvAcesso.value <> "" Then gerarArquivosJson opManifesto, Me.txtChvAcesso.value, Me.txtCaminhoPadrao.value
        MsgBox "Concluido!", vbInformation + vbOKOnly, strTitulo
        
        Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
        Me.Requery
        MsgBox "Concluido!", vbOKOnly, strTitulo
    Else
        MsgBox "Disponivel apenas para ""Registros valido - OK"" ", vbInformation + vbOKOnly, strTitulo
    End If
    
End Sub

Private Sub cmdManifestoTodos_Click()
    
    gerarArquivosJson opManifesto, , Me.txtCaminhoPadrao.value
    Shell "explorer " & Me.txtCaminhoPadrao.value, vbMaximizedFocus
    Me.Requery
    
End Sub

>>>>>>> f4084cb29d769387d25e7b837853d2119e0da429
