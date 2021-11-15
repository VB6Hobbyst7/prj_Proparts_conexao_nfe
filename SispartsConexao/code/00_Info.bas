Attribute VB_Name = "00_Info"
Option Compare Database

Private Const qryUpdateNumPed_CompraNF As String = "UPDATE TblCompraNF SET TblCompraNF.NumPed_CompraNF = Format(IIf(IsNull(DMax('NumPed_CompraNF', 'TblCompraNF')), '000001', DMax('NumPed_CompraNF', 'TblCompraNF') + 1), '000000') " & _
        "WHERE (((TblCompraNF.ID_CompraNF) IN (SELECT TOP 1 ID_CompraNF FROM TblCompraNF WHERE NumPed_CompraNF IS NULL ORDER BY ID_CompraNF)));"


''----------------------------

'' #AILTON - VALIDAR
'' #ARQUIVOS - GERAR ARQUIVOS | PROCESSAMENTO POR ARQUIVO(S)

'' #05_XML_ICMS                         - REVISÃO / FERNANDA
'' #05_XML_ICMS_Orig                    - REVISÃO / FERNANDA
'' #05_XML_ICMS_CST                     - REVISÃO / FERNANDA
'' #05_XML_ICMS_CST_VICMS               - REVISÃO / FERNANDA
'' #05_XML_IPI                          - REVISÃO / FERNANDA
'' #AILTON - qryInsertCompraItens       - REVISÃO / FERNANDA
'' #AILTON - qryInsertProdutoConsumo    - REVISÃO / FERNANDA
''
'' #PENDENTE - Processamento de arquivos CTE Inclusão de itens
'' #PENDENTE - Validação de campos de compras e itens
'' #PENDENTE - Teste de inclusão em banco SQL com todas as compras

''----------------------------

'' INFO 05/27/2021 17:10:29 - Processamento - Importar Dados Gerais ( Quantidade de registros: 1087 ) - 00:13:52
'' INFO 05/28/2021 15:19:27 - Processamento - Importar Registros Validos ( Quantidade de registros: 562 ) - 00:44:49


Public Function UpdateNumPed_CompraNF()
Dim x As Long
Dim contador As Long: contador = DCount("*", "TblCompraNF", "NumPed_CompraNF is null")

    For x = 1 To contador
        Application.CurrentDb.Execute qryUpdateNumPed_CompraNF
    Next

End Function
