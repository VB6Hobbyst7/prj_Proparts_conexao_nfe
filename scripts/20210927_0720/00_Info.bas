Attribute VB_Name = "00_Info"
Option Compare Database


'' #20210912_qryUpdateNumPed_CompraNF
'' #tblCompraNF.NumPed_CompraNF
Private Const qryUpdateNumPed_CompraNF As String = "UPDATE TblCompraNF SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000') " & _
        "WHERE (((TblCompraNF.ID_CompraNF) IN (SELECT TOP 1 ID_CompraNF FROM TblCompraNF WHERE NumPed_CompraNF IS NULL ORDER BY ID_CompraNF)));"


Private Const qryUpdateContador_NumPed_CompraNF As String = "UPDATE tblParametros SET tblParametros.ValorDoParametro = Format(IIf(IsNull(DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"")), '000001', DLookup(""[ValorDoParametro]"", ""[tblParametros]"", ""[TipoDeParametro]='NumPed_CompraNF'"") + 1), '000000') WHERE tblParametros.TipoDeParametro = ""NumPed_CompraNF"""


Sub azs__testeDeFuncionamentoGeral()

    '' LIMPAR BASE
    Application.CurrentDb.Execute "Delete from tblCompraNFItem"
    Application.CurrentDb.Execute "Delete from tblCompraNF"
    Application.CurrentDb.Execute "Delete from tblDadosConexaoNFeCTe"
    
    '' PROCESSAMENTO DE ARQUIVOS
    testeUnitario_carregarDadosGerais
    testeUnitario_carregarArquivosPendentes
    
    '' EXPORTAÇÃO DE DADOS
    testeUnitario_enviarDadosServidor
    
    '' TRATAMENTO DE ARQUIVOS
    testeUnitario_TratamentoDeArquivosValidos
    testeUnitario_TratamentoDeArquivosInvalidos
    
    Debug.Print "Concluido! - testeDeFuncionamentoGeral"

End Sub



'' #20210823_UpdateNumPed_CompraNF
Public Function UpdateNumPed_CompraNF()
Dim x As Long
Dim contador As Long: contador = DCount("*", "TblCompraNF", "NumPed_CompraNF is null")

Dim qryProcessos() As Variant: qryProcessos = Array(qryUpdateNumPed_CompraNF, qryUpdateContador_NumPed_CompraNF)

    For x = 1 To contador
        '' Application.CurrentDb.Execute qryUpdateNumPed_CompraNF
        executarComandos qryProcessos

    Next

End Function
