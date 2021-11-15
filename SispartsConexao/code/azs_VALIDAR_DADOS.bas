Attribute VB_Name = "azs_VALIDAR_DADOS"
Option Compare Database

Private Sub gerar_ArquivosDeValidacaoDeCampos()
Dim db As DAO.Database: Set db = CurrentDb
Dim rstRegistros As DAO.Recordset
Dim rstItens As DAO.Recordset
Dim sqlRegistros As String: sqlRegistros = "Select * from tblCompraNF where ChvAcesso_CompraNF = "
Dim sqlItens As String: sqlItens = "Select * from tblCompraNFItem where ChvAcesso_CompraNF = "
Dim arquivos As New Collection

Dim item As Variant
Dim TMP As String

'' 57
arquivos.add "42210300634453001303570010001139451001171544"

'' 55
'arquivos.add "32210368365501000296550000000639051001364146"

'' OUTROS
'arquivos.add "32210304884082000569570000040073831040073834"
'arquivos.add "42210220147617000494570010009539201999046070"
'arquivos.add "32210368365501000296550000000638811001361356"
'arquivos.add "42210212680452000302550020000886301507884230"
'arquivos.Add "32210368365501000296550000000638841001361501"


'' GERAR ARQUIVOS PARA VALIDAR DADOS
For Each item In arquivos


    '' Limpar repositorio de itens de compras
    Application.CurrentDb.Execute _
            "Delete from tblCompraNFItem where ChvAcesso_CompraNF = '" & CStr(item) & "'"

    '' Limpar repositorio de compras
    Application.CurrentDb.Execute _
            "Delete from tblCompraNF where ChvAcesso_CompraNF = '" & CStr(item) & "'"

    '' PROCESSAMENTO DE ARQUIVOS PENDENTES
    processarArquivosPendentes
    
    TMP = sqlRegistros & "'" & CStr(item) & "'"
    Set rstRegistros = db.OpenRecordset(TMP)
    
    Do While Not rstRegistros.EOF
        
        '' EXCLUIR ARQUIVO CASO EXISTA
        If (Dir(CurrentProject.path & "\" & CStr(item) & ".txt") <> "") Then Kill CurrentProject.path & "\" & CStr(item) & ".txt"
        
        '' CARREGAR CABEÇALHO
        TMP = ""
        For i = 0 To rstRegistros.Fields.count - 1
            TMP = rstRegistros.Fields(i).Name & vbTab & rstRegistros.Fields(i).value
            TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", TMP
        Next i

        '' CARREGAR ITENS DA NOTA
        TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", vbNewLine & "### ITENS ###" & vbNewLine

        TMP = ""
        TMP = sqlItens & "'" & CStr(item) & "'"
        Debug.Print TMP
        
        Set rstItens = db.OpenRecordset(TMP)
        Do While Not rstItens.EOF
            For i = 0 To rstItens.Fields.count - 1
                TMP = rstItens.Fields(i).Name & vbTab & rstItens.Fields(i).value
                TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", TMP
            Next i
            
            TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", vbNewLine & "#############################" & vbNewLine
            
            rstItens.MoveNext
            DoEvents
        Loop

        '' ABRIR ARQUIVO
        TMP = DLookup("[CaminhoDoArquivo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & CStr(item) & "'")
        Shell "notepad " & CurrentProject.path & "\" & CStr(item) & ".txt", vbMaximizedFocus
'        Debug.Print TMP
'        Shell "msedge.exe "" & TMP &"" ", vbMaximizedFocus
        
        Debug.Print "Concluido! - " & CStr(item) & ".txt"
        rstRegistros.MoveNext
        DoEvents
        TMP = ""
    Loop
    
    rstRegistros.Close
    rstItens.Close
Next

Debug.Print "Concluido!"
Set rstRegistros = Nothing
Set rstItens = Nothing

End Sub
