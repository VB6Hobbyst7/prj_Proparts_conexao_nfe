Attribute VB_Name = "basConexao"
Option Compare Database
Public CNN As ADODB.Connection

Function AbrirConexao()
On Error GoTo ConexaoTrataErro
Dim ObjConnErr As Boolean
Dim strConnect As String
'esta é a string de conexao devera conter a informacao sobre o provedor e o caminho do banco de dados
Dim strProvider As String
'guarda o nome do provedor
Dim strDataSource As String
'guarda a fonte de dados
Dim strDataBaseName As String

'nome do banco de dados
'Dim usr_id As String       ' identificacao do usuario para o banco de dados
'Dim pass As String         ' a senha (se tiver) para o banco de dados
'Dim mySqlIP As String    ' o endereco/ip da maquina na qual esta o mySql
'mySqlIP = "10.0.2.173"    ' a localizacao do usuario (localhost)
'mySqlIP = "servidor"    ' a localizacao do usuario (localhost) - Arte Micro
'mySqlIP = "192.168.0.10"    ' a localizacao do usuario (linux) - Arte Micro
'mySqlIP = "MICRO5"    ' a localizacao do usuario (localhost) - Micro5
'usr_id = "root"  ' identificacao
'pass = ""       ' senha
' string de conexao
'strConnect = "driver={MySQL ODBC 3.51 Driver};server=" & mySqlIP & ";uid=" & usr_id & ";pwd=" & pass & ";database=PML"

ObjConnErr = True
If CNN.State = 1 Then
 If ObjConnErr Then Exit Function
End If
ObjConnErr = False

CurrentDb.QueryDefs.Delete "GetActiveConnection"
strConnect = CurrentDb.CreateQueryDef("GetActiveConnection", "SELECT String_Con as Conexao FROM tblConexao WHERE Tipo_Con='SQL' AND FlagPadrao_Con=-1 ;").OpenRecordset!Conexao
CurrentDb.QueryDefs.Delete "GetActiveConnection"

Set CNN = New ADODB.Connection

'preparando o objeto connection
CNN.CursorLocation = adUseClient

'usamos um cursor do lado do cliente pois os dados serao acessados na maquina do cliente e nao de um servidor
CNN.Open strConnect

'Abre o objeto connection

Exit Function

ConexaoTrataErro:
If err.Number = 91 Then ObjConnErr = False: Resume Next
If err.Number = 3265 Then Resume Next
  For Each adoErro In CNN.Errors
    MsgBox adoErro.description
  Next
End Function
Public Function GetState(intState As Integer) As String

    Select Case intState
        Case adStateClosed
            GetState = "adStateClosed"
        Case adStateOpen
            GetState = "adStateOpen"
    End Select

End Function

Public Function TestarConexao(StringConexao As String) As Boolean
On Error GoTo NoBug
Dim ConnTest As New ADODB.Connection

'SEMPRE DEIXAR ESTE VALOR COMO TRUE
TestarConexao = True

'Abrindo a conexão para verificar se está correta.
ConnTest.Open StringConexao
ConnTest.Close

Exit Function
NoBug:
'Caso ocorra algum erro, a conexão não é válida
TestarConexao = False
Exit Function
End Function

