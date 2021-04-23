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

'' #carregarCompras
'' PROCESSAMENTO DAS COMPRAS COM BASE EM REGISTROS VALIDOS PROCESSADOS PELA #CAPTURA_DADOS_GERAIS
Public Const qrySelectProcessamentoPendente As String = _
            "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo, tblDadosConexaoNFeCTe.ID_Tipo FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.ID_Tipo)>0))"

'Public Const qryUpdateProcessamentoConcluido As String = _
'            "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) " & _
'            " AND ((tblDadosConexaoNFeCTe.registroProcessado)=0) AND ((tblDadosConexaoNFeCTe.Chave)='strChave'));"

Public Enum enumOperacao
    opNone = 0
    opExecutar = 1
    opNotepad = 2
End Enum


'' #####################################################################
'' ### #Libs - PODE SER ADICIONADAS AS FUNÇÕES GERAIS DA APLICAÇÃO
'' #####################################################################

Public Function getConsultarSeRetornoArmazemParaRecuperarNumeroDePedido(pChvAcesso As String, pDados As String) As String
'' valor padrao
Dim tRetorno As String: tRetorno = 0

'' Codigo do Retorno de armazem
Dim tTipo As String: tTipo = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='RetornoArmazem'")

'' tipo de cadastro
Dim tTipoCadastro As String: tTipoCadastro = DLookup("[ID_Tipo]", "[tblDadosConexaoNFeCTe]", "[ChvAcesso]='" & pChvAcesso & "'")

'' limpar dado inicial
Dim tValor() As Variant: tValor = Array("PEDIDO:", "PEDIDO")

    
    If tTipoCadastro = tTipo Then
        tRetorno = left(Trim(Replace(Replace(pDados, tValor(0), ""), tValor(1), "")), 6)
    End If
    
    
    getConsultarSeRetornoArmazemParaRecuperarNumeroDePedido = tRetorno

End Function


'' #VALIDAR_DADOS
Sub criarConsultasParaTestes()
Dim db As dao.Database: Set db = CurrentDb
Dim rstOrigem As dao.Recordset
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
    qryDeleteExists "qry_" & tabela
    qryCreate "qry_" & tabela, strSql
Next tabela

db.Close: Set db = Nothing

End Sub

Sub ImportarDados(pOrigem As String, pDestino As String)

'' #BANCO_ORIGEM
Dim strUsuarioNome As String: strUsuarioNome = pegarValorDoParametro(qryParametros, "BancoDados_Usuario")
Dim strUsuarioSenha As String: strUsuarioSenha = pegarValorDoParametro(qryParametros, "BancoDados_Senha")
Dim strOrigem As String: strOrigem = pegarValorDoParametro(qryParametros, "BancoDados_Origem")
Dim strBanco As String: strBanco = pegarValorDoParametro(qryParametros, "BancoDados_Banco")
Dim tmpOrigem As String: tmpOrigem = "Select * from " & pOrigem

Dim dboOrigem As New Banco: dboOrigem.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
dboOrigem.SqlSelect tmpOrigem

'' #BANCO_LOCAL
Dim db As dao.Database: Set db = CurrentDb
Dim tmpDestino As String: tmpDestino = "Select * from " & pDestino
Dim rstDestino As dao.Recordset: Set rstDestino = db.OpenRecordset(tmpDestino)
Dim rstOrigem As dao.Recordset

'' #ANALISE_DE_PROCESSAMENTO
Dim DT_PROCESSO As Date: DT_PROCESSO = Now()

'' #BARRA_PROGRESSO
Dim contadorDeRegistros As Long: contadorDeRegistros = 1
SysCmd acSysCmdInitMeter, "Transferindo " & pOrigem & " ...", dboOrigem.rs.RecordCount

Do While Not dboOrigem.rs.EOF

    '' #BARRA_PROGRESSO
    SysCmd acSysCmdUpdateMeter, contadorDeRegistros

    '' listar campos da tabela
    For Each fld In rstDestino.Fields
        If t = 0 Then rstDestino.AddNew: t = 1
        rstDestino.Fields(fld.Name).value = dboOrigem.rs.Fields(fld.Name).value
        DoEvents
    Next
    rstDestino.Update
    t = 0
    
    '' #BARRA_PROGRESSO
    contadorDeRegistros = contadorDeRegistros + 1
    dboOrigem.rs.MoveNext
    DoEvents
Loop

'' #BARRA_PROGRESSO
SysCmd acSysCmdRemoveMeter

'dbDestino.CloseConnection
db.Close: Set db = Nothing

'' #ANALISE_DE_PROCESSAMENTO
statusFinal DT_PROCESSO, "Processamento - ImportarDados ( " & pDestino & " )"

End Sub

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

'' EXECUÇÃO DE APLICATIVO EXTERNO
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
    
    '' Saida para arquivo
    TextFile_Append strFilePath & strFileName, tmp
    
    '' Operações - Execução ou Abrir arquivo depois de pronto
    Dim pathApp As String: pathApp = strApp & " " & strFilePath & strFileName
    Select Case pOperacao
        Case opExecutar
            Shell pathApp
        Case opNotepad
            Shell pathApp, vbMaximizedFocus
        Case Else
    End Select
    
End Function

'' CRIACAO DE ARQUIVOS
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

'' EXECUTAR CONSULTAS
Public Sub executarComandos(comandos() As Variant) ''#ExecutarConsultas
Dim Comando As Variant

    For Each Comando In comandos
        Application.CurrentDb.Execute Comando
    Next Comando

End Sub

'' CRIACAO DE CONSULTAS EM TEMPO DE EXECUCAO
Public Sub qryCreate(nomeDaConsulta As String, scriptDaConsulta As String)
Dim db As dao.Database: Set db = CurrentDb
    db.CreateQueryDef nomeDaConsulta, scriptDaConsulta
    db.Close
End Sub

'' VERIFICAR A EXISTENCIA DE CONSULTAS E EXCLUI CASO EXISTA
Public Function qryDeleteExists(strQryName As String)
Dim db As dao.Database: Dim qdf As dao.QueryDef
    For Each qdf In CurrentDb.QueryDefs
        If qdf.Name = strQryName Then
            CurrentDb.QueryDefs.Delete strQryName
            Exit For
        End If
    Next
End Function

'' CARREGAR PARAMETROS COMPOSTOS
Public Function carregarParametros(pConsulta As String, Optional pParametro As String) As Collection: Set carregarParametros = New Collection
Dim db As dao.Database: Set db = CurrentDb
Dim strSql As String: strSql = IIf(pParametro <> "", Replace(pConsulta, "strParametro", pParametro), pConsulta)
Dim rst As dao.Recordset: Set rst = db.OpenRecordset(strSql)
Dim f As Variant

Do While Not rst.EOF
    carregarParametros.add rst.Fields(0).value
    rst.MoveNext
Loop

db.Close

Set db = Nothing

End Function

'' CARREGAR PARAMETROS UNICOS
Public Function pegarValorDoParametro(pConsulta As String, pTipoDeParametro As String, Optional pCampo As String) As String
Dim db As dao.Database: Set db = CurrentDb
Dim strTmp As String: strTmp = Replace(pConsulta, "strParametro", pTipoDeParametro)
Dim rst As dao.Recordset: Set rst = db.OpenRecordset(strTmp)

    pegarValorDoParametro = rst.Fields(IIf(pCampo <> "", pCampo, "ValorDoParametro")).value

db.Close
End Function

''#ExtrairCaminhoDoArquivo
Public Function getPath(sPathIn As String) As String
Dim i As Integer

  For i = Len(sPathIn) To 1 Step -1
     If InStr(":\", Mid$(sPathIn, i, 1)) Then Exit For
  Next

  getPath = left$(sPathIn, i)

End Function

''#ExtrairNomeDoArquivoComExtensao
Public Function getFileNameAndExt(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquiva
Dim i As Integer

    For i = Len(sFileIn) To 1 Step -1
       If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    
    getFileNameAndExt = Mid$(sFileIn, i + 1, Len(sFileIn) - i)

End Function

''#ExtrairNomeDoArquivoSemExtensao
Public Function getFileName(sFileIn As String) As String
' Essa função irá retornar apenas o nome do  arquivo de uma
' string que contenha o path e o nome do arquivo
Dim i As Integer

  For i = Len(sFileIn) To 1 Step -1
     If InStr("\", Mid$(sFileIn, i, 1)) Then Exit For
  Next

  getFileName = left(Mid$(sFileIn, i + 1, Len(sFileIn) - i), Len(Mid$(sFileIn, i + 1, Len(sFileIn) - i)) - 4)

End Function

'' LIMPAR COLEÇOES
Public Sub ClearCollection(ByRef container As Collection)
    Dim index As Long
    For index = 1 To container.Count
        container.remove 1
    Next
End Sub

'' STATUS DE PROCESSAMENTO
Public Function statusFinal(pDate As Date, strTitulo As String)
Dim tmp As String: tmp = "INFO " & Format(Now, "mm/dd/yyyy HH:mm:ss") & " - " & strTitulo & " - " & Format(Now - pDate, "hh:mm:ss")
Dim strFileName As String: strFileName = right(Year(Now()), 4) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & ".log"
    
    '' Saida de Log
    TextFile_Append CurrentProject.path & "\" & strFileName, tmp
    
End Function

'' SEPARAR DADOS COMPOSTOS
Public Function strSplit(strValor As String, strSeparador As String, intPosicao As Integer) As String
    If (strValor <> "") Then
        strSplit = Split(strValor, strSeparador)(intPosicao)
    Else
        strSplit = ""
    End If
End Function

'' GERAR IDENTIFICADOR UNICO
Public Function Controle() As String
    Controle = right(Year(Now()), 4) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & Format(Hour(Now()), "00") & Format(Minute(Now()), "00") & Format(Second(Now()), "00")
End Function

'' GERAR IDENTIFICADOR PARA CRIAÇÃO DE PASTA
Public Function strControle() As String
    strControle = right(Year(Now()), 4) & Format(Month(Now()), "00") & Format(Day(Now()), "00") & "_" & Format(Hour(Now()), "00") & Format(Minute(Now()), "00")
End Function

'' LIMPAR PONTOS
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

Public Function NumberToSql(pNumber As Variant) As String
    Dim i As Integer
    Dim iStrSinal As String
    Dim iStrAux As String
    Dim iCaracteres As String
    iCaracteres = "1234567890,.-"
    iStrAux = vbNullString
    iStrSinal = vbNullString
    
    pNumber = CStr(pNumber)
    
    If pNumber = vbNullString Then
        NumberToSql = "NULL"
    Else
        For i = 1 To Len(pNumber)
            If Mid(pNumber, i, 1) = "(" Then iStrSinal = "-"
            If InStr(iCaracteres, Mid(pNumber, i, 1)) <> 0 Then
                iStrAux = iStrAux & Mid(pNumber, i, 1)
            End If
        Next i
        iStrAux = Replace(iStrAux, Chr(46), vbNullString)       ' Remove pontos
        iStrAux = Replace(iStrAux, Chr(44), Chr(46))            ' Substitui vírgula por ponto
        NumberToSql = iStrSinal & iStrAux
        If right(NumberToSql, 1) = "." Then NumberToSql = left(NumberToSql, Len(NumberToSql) - 1)

        If Not IsNumeric(NumberToSql) Then
            Debug.Print Now & " -:- ERRO Function NumberToSql -:- Impossível formatar " & pNumber & ". Resultado não é um número válido: " & NumberToSql
            NumberToSql = "0"
        End If
    End If
End Function

Function createTable(pTabelaDestino As String) As String

Dim db As dao.Database: Set db = CurrentDb
Dim rstCampos As dao.Recordset: Set rstCampos = db.OpenRecordset("Select distinct campo from tblOrigemDestino where tabela = '" & pTabelaDestino & "' and len(formatacao)>0 order by campo;")

Dim tmpScript As String: tmpScript = "CREATE TABLE " & pTabelaDestino & " ("

Do While Not rstCampos.EOF
    tmpScript = tmpScript + rstCampos.Fields("campo").value + " text(255),"
    rstCampos.MoveNext
    DoEvents
Loop

createTable = left(tmpScript, Len(tmpScript) - 1) + ");"

End Function


Public Function PickFolder(pPath As String, pTitle As String, Optional pSubFolder As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = pTitle
        .InitialFileName = pPath
        .Show
            If .SelectedItems.Count > 0 Then
                If Trim(pSubFolder) <> "" Then
                    If Dir(.SelectedItems(1) & pSubFolder, vbDirectory) = "" Then
                        MkDir path:=.SelectedItems(1) & pSubFolder
                    End If
                End If
                PickFolder = .SelectedItems(1) & pSubFolder
            End If
    End With
End Function
