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

