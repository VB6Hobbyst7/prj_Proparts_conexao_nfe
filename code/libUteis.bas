Attribute VB_Name = "libUteis"
Option Compare Database
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'' #CONTROLE_PARAMETRO
Public Const qryParametros As String = "SELECT tblParametros.ValorDoParametro FROM tblParametros WHERE (((tblParametros.TipoDeParametro) = 'strParametro'))"
Public Const strCaminhoDeColeta As String = "caminhoDeColeta"
Public Const strCaminhoDeProcessados As String = "caminhoDeProcessados"
Public Const strUsuarioErpCod As String = "UsuarioErpCod"
Public Const strUsuarioErpNome As String = "UsuarioErpNome"
Public Const strCodTipoEvento As String = "codTipoEvento"
Public Const strComando As String = "Comando"

'' #CAPTURA_COMPRAS
'' PROCESSAMENTO DAS COMPRAS COM BASE EM REGISTROS VALIDOS PROCESSADOS PELA #CAPTURA_DADOS_GERAIS
Public Const qrySelectProcessamentoPendente As String = "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));"
Public Const qryUpdateProcessamentoConcluido As String = "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.Chave)='strChave'));"


'' COMPRAS
'' -- AJUSTE DE CAMPOS
Private Const qryUpdateCurrecy As String = "UPDATE tblProcessamento SET tblProcessamento.valor = Format(Replace([tblProcessamento].[valor], '.', ','), '#,##0.00') WHERE (((tblProcessamento.NomeCampo)='strCampo'));"
Private Const qryUpdateDate As String = "UPDATE tblProcessamento SET tblProcessamento.valor = CDate(Replace(Mid([tblProcessamento].[valor],1,10),'-','/')) WHERE (((tblProcessamento.NomeCampo)='strCampo'));"
Private Const qryUpdateTime As String = "UPDATE tblProcessamento SET tblProcessamento.valor = Replace(Mid([tblProcessamento].[valor],12,8),'-','/') WHERE (((tblProcessamento.NomeCampo)='strCampo'));"


Public Enum enumOperacao
    opNone = 0
    opExecutar = 1
    opNotepad = 2
End Enum


'' #####################################################################
'' ### #Libs - PODE SER ADICIONADAS AS FUNÇÕES GERAIS DA APLICAÇÃO
'' #####################################################################

Public Function statusFinal(pDate As Date, strTitulo As String)
    
    Debug.Print strTitulo & " - " & Format(Now - pDate, "hh:mm:ss")
    
End Function



'' #VALIDAR_DADOS
Sub criarConsultasParaTestes()
Dim db As DAO.Database: Set db = CurrentDb
Dim rstOrigem As DAO.Recordset
Dim strSql As String
Dim qrySelectTabelas As String: qrySelectTabelas = "Select Distinct tabela from tblOrigemDestino order by tabela"
Dim tabela As Variant

'' CRIAR CONSULTA PARA VALIDAR DADOS PROCESSADOS
For Each tabela In carregarParametros(qrySelectTabelas)
    strSql = "Select "
    Set rstOrigem = db.OpenRecordset("Select distinct Destino from tblOrigemDestino where tabela = '" & tabela & "'")
    Do While Not rstOrigem.EOF
        strSql = strSql & strSplit(rstOrigem.Fields("Destino").value, ".", 1) & ","
        rstOrigem.MoveNext
    Loop
    strSql = left(strSql, Len(strSql) - 1) & " from " & tabela
    qryExists "qry_" & tabela
    qryCreate "qry_" & tabela, strSql
Next tabela

db.Close: Set db = Nothing

End Sub

'' #FORMATAR_CAMPOS
Sub formatarCampos()
Dim t As Variant
Dim s As String

'' MOEDA
For Each t In Array("BaseCalcICMSSubsTrib_CompraNF", "BaseCalcICMS_CompraNF", "VTotICMS_CompraNF", "VTotServ_CompraNF", "VTotProd_CompraNF", "VTotNF_CompraNF", "VTotICMSSubsTrib_CompraNF", "VTotFrete_CompraNF", "VTotSeguro_CompraNF", "VTotOutDesp_CompraNF", "VTotIPI_CompraNF", "VTotISS_CompraNF", "TxDesc_CompraNF", "VTotDesc_CompraNF")
    Application.CurrentDb.Execute Replace(qryUpdateCurrecy, "strCampo", t)
Next t

'' DATAS
For Each t In Array("DTEmi_CompraNF", "DTEntd_CompraNF")
    Application.CurrentDb.Execute Replace(qryUpdateDate, "strCampo", t)
Next t

'' HORAS
For Each t In Array("HoraEntd_CompraNF")
    Application.CurrentDb.Execute Replace(qryUpdateTime, "strCampo", t)
Next t

End Sub


Public Function strSplit(strValor As String, strSeparador As String, intPosicao As Integer) As String
    If (strValor <> "") Then
        strSplit = Split(strValor, strSeparador)(intPosicao)
    Else
        strSplit = ""
    End If
End Function

Public Function Controle() As String
    Controle = right(Year(Now()), 2) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & Format(Hour(Now()), "00") & Format(Minute(Now()), "00") & Format(Second(Now()), "00")
End Function

Public Function GetFilesInSubFolders(pFolder As String) As Collection
Set GetFilesInSubFolders = New Collection

    Dim objFSO As Object: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder As Object: Set objFolder = objFSO.GetFolder(pFolder)
    Dim objSubFolders As Object
    Dim objFile As Object
    Dim iCol As New Collection
            
    For Each objSubFolders In objFolder.subFolders
        For Each objFile In objSubFolders.files
            GetFilesInSubFolders.add objFile.path
        Next objFile
    Next objSubFolders
    
    Set objFSO = Nothing
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objSubFolders = Nothing
    
'    Set GetFilesInSubFolders = iCol
End Function

Public Function CreateDir(strPath As String) As String
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
    
    CreateDir = strPath
    
End Function


Public Function execucao(pCol As Collection, strFileName As String, Optional strFilePath As String, Optional pOperacao As enumOperacao, Optional strApp As String) 'runUrl.au3
Dim c As Variant, tmp As String: tmp = ""
    
    '' Path
    If ((strFilePath) = "") Then
        strFilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    Else
        strFilePath = strFilePath & "\"
    End If
    
    If (Dir(strFilePath & strFileName) <> "") Then Kill strFilePath & strFileName
    
    '' Criação
    For Each c In pCol
        tmp = tmp + CStr(c) + vbNewLine
    Next c
    TextFile_Append strFilePath & strFileName, tmp
    
    Dim pathApp As String: pathApp = strApp & " " & strFilePath & strFileName
    Select Case pOperacao
        Case opExecutar
            Shell pathApp
        Case opNotepad
            'ClipBoardThis tmp
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

Public Sub qryCreate(nomeDaConsulta As String, scriptDaConsulta As String)
Dim db As DAO.Database: Set db = CurrentDb
    db.CreateQueryDef nomeDaConsulta, scriptDaConsulta
    db.Close
End Sub

Public Function qryExists(strQryName As String)
Dim db As DAO.Database: Dim qdf As DAO.QueryDef
    For Each qdf In CurrentDb.QueryDefs
        If qdf.Name = strQryName Then
            CurrentDb.QueryDefs.Delete strQryName
            Exit For
        End If
    Next
End Function

Public Function carregarParametros(pConsulta As String, Optional pParametro As String) As Collection: Set carregarParametros = New Collection
Dim db As DAO.Database: Set db = CurrentDb
Dim strSql As String: strSql = IIf(pParametro <> "", Replace(pConsulta, "strParametro", pParametro), pConsulta)
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(strSql)
Dim f As Variant

Do While Not rst.EOF
    carregarParametros.add rst.Fields(0).value
    rst.MoveNext
Loop

db.Close

Set db = Nothing

End Function

Public Function pegarValorDoParametro(pConsulta As String, pTipoDeParametro As String, Optional pCampo As String) As String
Dim db As DAO.Database: Set db = CurrentDb
Dim rst As DAO.Recordset: Set rst = db.OpenRecordset(Replace(pConsulta, "strParametro", pTipoDeParametro))

    pegarValorDoParametro = rst.Fields(IIf(pCampo <> "", pCampo, "ValorDoParametro")).value

db.Close
End Function

Public Function getPath(sPathIn As String) As String ''#ExtrairCaminhoDoArquivo
Dim i As Integer

  For i = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
  Next

  getPath = left$(sPathIn, i)

End Function

Public Function getFileNameAndExt(sFileIn As String) As String ''#ExtrairNomeDoArquivo
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

    For i = Len(sFileIn) To 1 Step -1
       If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    
    getFileNameAndExt = Mid$(sFileIn, i + 1, Len(sFileIn) - i)

End Function

Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

  For i = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
  Next

  getFileName = left(Mid$(sFileIn, i + 1, Len(sFileIn) - i), Len(Mid$(sFileIn, i + 1, Len(sFileIn) - i)) - 4)

End Function

Public Sub ClearCollection(ByRef container As Collection)
    Dim index As Long
    For index = 1 To container.Count
        container.remove 1
    Next
End Sub

Public Function STRPontos(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim nova As String
  Dim x
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case ".", ",", "-", " ", "/", "\"
        x = ""
      Case Else
      x = UCase$(x)
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRPontos = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Public Function PreventNullString(pText As Variant) As String
    If IsNull(pText) Then
        PreventNullString = ""
    Else
        PreventNullString = CStr(pText)
    End If
End Function
