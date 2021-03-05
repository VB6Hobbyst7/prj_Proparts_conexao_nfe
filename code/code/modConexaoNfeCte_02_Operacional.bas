Attribute VB_Name = "modConexaoNfeCte_02_Operacional"
Option Compare Database

Private FileCollention As New Collection
Private con As ADODB.Connection

Public Enum enumOperacao
    opNome = 0
    opExecutar = 1
    opNotepad = 2
End Enum

Private Const script_tblDadosConexaoNFeCTe As String = "INSERT INTO tblDadosConexaoNFeCTe (codMod,dhEmi,CNPJ_emit,Razao_emit,CNPJ_Rem,CPNJ_Dest,CaminhoDoArquivo) VALUES(strOrigem,'strCaminhoDoArquivo')"
Private Const parametro As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strTipoDeParametro'))"
Private Const strCaminhoDeColeta As String = "caminhoDeColeta"


'' #Libs
'' carregarDados
'' carregarERP
'' carregarManifesto
'' main ---->   EM DESENVOLVIMENTO

''#Ailton - EM TESTES
Sub teste_carregarDados()
Dim tmp As clsDadosConexaoNFeCTe: Set tmp = New clsDadosConexaoNFeCTe

    tmp.cadastroDados pegarValorDoParametro(strCaminhoDeColeta)

End Sub

'' #####################################################################
'' ### #Libs - PODE SER ADICIONADAS AS FUNÇÕES GERAIS DA APLICAÇÃO
'' #####################################################################

Public Function GetFilesInSubFolders(pFolder As String) As Collection
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

Public Function CreateDir(strPath As String)
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Function

Public Function execucao(pCol As Collection, strFileName As String, Optional strFilePath As String, Optional pOperacao As enumOperacao, Optional strApp As String) 'runUrl.au3
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

Public Function TextFile_Append(pFilePath As String, pText As String) As Boolean
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

Sub executarComandos(comandos() As Variant)
Dim Comando As Variant

    For Each Comando In comandos
        Application.CurrentDb.Execute Comando
    Next Comando

End Sub

Sub CadastroDeItens(Itens As Collection)
Dim con As ADODB.Connection: Set con = CurrentProject.Connection
Dim i As Variant

    For Each i In Itens
        con.Execute i
    Next i

Set con = Nothing

End Sub

Sub criarConsulta(nomeDaConsulta As String, scriptDaConsulta As String)
Dim db As DAO.Database: Set db = CurrentDb

    db.CreateQueryDef nomeDaConsulta, scriptDaConsulta
    db.Close

End Sub

Function pegarValorDoParametro(pTipoDeParametro As String) As String
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(Replace(parametro, "strTipoDeParametro", pTipoDeParametro))

    pegarValorDoParametro = rst.Fields("ValorDoParametro").Value

End Function

Function timestamp() As Double
Dim iUnixTime As Long: iUnixTime = Now()
    timestamp = ((((iUnixTime / 60) / 60) / 24) + 255690416666667#)
End Function

Public Function getPath(sPathIn As String) As String
Dim i As Integer

  For i = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
  Next

  getPath = left$(sPathIn, i)

End Function

Public Function getFileNameAndExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

    For i = Len(sFileIn) To 1 Step -1
       If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    
    getFileNameAndExt = Mid$(sFileIn, i + 1, Len(sFileIn) - i)

End Function


Sub LeituraDeArquivos() '' Leitura de arquivo para tabela de dados
    Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
    
    '' 01. Leitura e identificação do arquivo
    For Each fileName In GetFilesInSubFolders(pegarValorDoParametro(strCaminhoDeColeta))

        XDoc.Load fileName
        
        Debug.Print vbNewLine
        Debug.Print fileName
        
        '' 02. Separação por tags e coleta de dados
        For Each Item In Array("ide/mod", "ide/dhEmi", "emit/CNPJ", "emit/xNome", "rem/CNPJ", "dest/CNPJ", "infCTeNorm/infDoc/infNFe/chave")
            Set lists = XDoc.SelectNodes("//" & Item)
            For Each fieldNode In lists
                If (fieldNode.HasChildNodes) Then
                    For Each childNode In fieldNode.ChildNodes
                        Debug.Print "[" & fieldNode.BaseName & "] = [" & fieldNode.text & "]"
                    Next childNode
                End If
            Next fieldNode
        Next

    Next

    Set XDoc = Nothing
End Sub
