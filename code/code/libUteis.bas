Attribute VB_Name = "libUteis"
Option Compare Database

Public Enum enumOperacao
    opNome = 0
    opExecutar = 1
    opNotepad = 2
End Enum


'' #####################################################################
'' ### #Libs - PODE SER ADICIONADAS AS FUNÇÕES GERAIS DA APLICAÇÃO
'' #####################################################################

Public Function Controle() As String
    Controle = right(Year(Now()), 2) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & Format(Hour(Now()), "00") & Format(Minute(Now()), "00") & Format(Second(Now()), "00")
End Function

Public Function GetFilesInSubFolders(pFolder As String) As Collection ''#ConsultarArquivosEmPastas
    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder As Object: Set objFolder = objFSO.GetFolder(pFolder)
    Dim objSubFolders As Object
    Dim objFile As Object
    Dim iCol As New Collection
            
    For Each objSubFolders In objFolder.subFolders
        For Each objFile In objSubFolders.files
            iCol.add objFile.path
        Next objFile
    Next objSubFolders
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objSubFolders = Nothing
    
    Set GetFilesInSubFolders = iCol
End Function

Public Function CreateDir(strPath As String) ''#CriarPastasDestino
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Function

Public Function execucao(pCol As Collection, strFileName As String, Optional strFilePath As String, Optional pOperacao As enumOperacao, Optional strApp As String) ''#CriarArquivosJson
Dim c As Variant, tmp As String: tmp = ""
    
    '' Seleção do diretorio
    If ((strFilePath) = "") Then
        strFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    Else
        strFilePath = strFilePath & "\"
    End If
        
    '' Criação do diretorio caso não exista no caminho solicitado
    If (Dir(strFilePath & strFileName) <> "") Then Kill strFilePath & strFileName
    
    '' Criação do arquivo
    For Each c In pCol
        tmp = tmp + CStr(c) + vbNewLine
    Next c
    TextFile_Append strFilePath & strFileName, tmp
    
    '' Execução do arquivo criado
    Dim pathApp As String: pathApp = strApp & " " & strFilePath & strFileName
    Select Case pOperacao
        Case opExecutar
            Shell pathApp
        Case opNotepad
            Shell pathApp, vbMaximizedFocus
        Case Else
    End Select
    
End Function

Public Function TextFile_Append(pFilePath As String, pText As String) As Boolean ''#CriarArquivosJson (Auxiliar)
    Dim intFNumber As Integer
    intFNumber = FreeFile
    
    If boolErrorHandler Then On Error GoTo ErrorHandler
    
    Open pFilePath For Append As #intFNumber
    Print #intFNumber, pText
    Close #intFNumber
    
    On Error GoTo 0
    
    TextFile_Append = True
    Exit Function
ErrorHandler:
    MsgBox "Não foi possível salvar o arquivo." & vbNewLine & "Verifique o caminho informado e as permissões de acesso", vbInformation
End Function

Public Sub executarComandos(comandos() As Variant) ''#ExecutarConsultas
Dim Comando As Variant

    For Each Comando In comandos
        Application.CurrentDb.Execute Comando
    Next Comando

End Sub

Public Function qryExists(strQryName As String) As Boolean: qryExists = False
Dim db As DAO.Database
Dim qdf As DAO.QueryDef

For Each qdf In CurrentDb.QueryDefs
    If qdf.Name = strQryName Then
        qryExists = True
        Exit For
    End If
Next

End Function

Public Function carregarParametros(pParametro As String, pConsulta As String) As Collection
Set carregarParametros = New Collection
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(Replace(pConsulta, "strParametro", pParametro))
Dim f As Variant

Do While Not rst.EOF
    carregarParametros.add rst.Fields(0).Value
    rst.MoveNext
Loop

db.Close

Set db = Nothing

End Function

Public Function pegarValorDoParametro(pTipoDeParametro As String, pConsulta As String) As String ''#CarregarValorDeParametro
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(Replace(pConsulta, "strParametro", pTipoDeParametro))

    pegarValorDoParametro = rst.Fields("ValorDoParametro").Value

End Function

Public Function getPath(sPathIn As String) As String ''#ExtrairCaminhoDoArquivo
Dim I As Integer

  For I = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, I, 1)) Then Exit For
  Next

  getPath = left$(sPathIn, I)

End Function

Public Function getFileNameAndExt(sFileIn As String) As String ''#ExtrairNomeDoArquivo
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

    For I = Len(sFileIn) To 1 Step -1
       If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
    Next
    
    getFileNameAndExt = Mid$(sFileIn, I + 1, Len(sFileIn) - I)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim I As Integer

  For I = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, I, 1)) Then Exit For
  Next

  getFileName = left(Mid$(sFileIn, I + 1, Len(sFileIn) - I), Len(Mid$(sFileIn, I + 1, Len(sFileIn) - I)) - 4)

End Function

Public Sub ClearCollection(ByRef container As Collection)
    Dim index As Long
    For index = 1 To container.count
        container.remove 1
    Next
End Sub
