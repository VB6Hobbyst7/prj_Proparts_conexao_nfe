Attribute VB_Name = "azs_VALIDAR_DADOS"
Option Compare Database


Sub teste_carregar()

Dim pChave As String: pChave = "42210348740351012767570000021186731952977908-cteproc"
Dim pRepositorio As String: pRepositorio = "tblCompraNF"

Dim tmpValidarCampo As String: tmpValidarCampo = right(pRepositorio, Len(pRepositorio) - 3)
Dim tmpScript As String: tmpScript = Replace("SELECT NomeCampo,valor,formatacao FROM tblProcessamento where pk = 'pChave' and not NomeCampo is null", "pChave", pChave)
Dim tmpScriptCabecalho As String
Dim item As Variant

Dim db As DAO.Database: Set db = CurrentDb
Dim rstRepositorio As DAO.Recordset: Set rstRepositorio = db.OpenRecordset(tmpScript)
'Debug.Print tmpScript

tmpScript = ""
Do While Not rstRepositorio.EOF

    For Each item In Split(carregarCamposNomes(pRepositorio), ",")
        
        If InStr(rstRepositorio.Fields("NomeCampo").value, CStr(item)) Then
        
            If rstRepositorio.Fields("formatacao").value = "opTexto" Then
                tmpScript = tmpScript & "'" & rstRepositorio.Fields("Valor").value & "',"
        
            ElseIf rstRepositorio.Fields("formatacao").value = "opNumero" Or rstRepositorio.Fields("formatacao").value = "opMoeda" Then
                tmpScript = tmpScript & rstRepositorio.Fields("Valor").value & ","
        
            ElseIf rstRepositorio.Fields("formatacao").value = "opTime" Then
                tmpScript = tmpScript & "'" & Format(rstRepositorio.Fields("Valor").value, DATE_TIME_FORMAT) & "',"
        
            ElseIf rstRepositorio.Fields("formatacao").value = "opData" Then
                tmpScript = tmpScript & "'" & Format(rstRepositorio.Fields("Valor").value, DATE_FORMAT) & "',"
        
            End If
        
            tmpScriptCabecalho = tmpScriptCabecalho & rstRepositorio.Fields("NomeCampo").value & ","
        End If
        DoEvents
    Next
        
    rstRepositorio.MoveNext
    DoEvents
Loop

Debug.Print left(tmpScriptCabecalho, Len(tmpScriptCabecalho) - 1)
Debug.Print left(tmpScript, Len(tmpScript) - 1)

End Sub
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
