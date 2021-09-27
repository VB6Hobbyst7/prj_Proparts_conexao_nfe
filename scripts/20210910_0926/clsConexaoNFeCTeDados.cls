VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConexaoNFeCTeDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private pCol As New Collection

Private Type ClassType
        dhEmi As String
        Comando As String
        codTipoEvento As String
        ChvAcesso As String
        codUsuarioErp As String
        nomeUsuarioErp As String
        caminhoDeProcessados As String
        strFileName As String
End Type
Private This As ClassType

Property Get dhEmi() As String: dhEmi = This.dhEmi: End Property
Property Let dhEmi(pValue As String): This.dhEmi = pValue: End Property

Property Get ChvAcesso() As String: ChvAcesso = This.ChvAcesso: End Property
Property Let ChvAcesso(pValue As String): This.ChvAcesso = pValue: End Property

Property Get Comando() As String: Comando = This.Comando: End Property
Property Let Comando(pValue As String): This.Comando = pValue: End Property

Property Get codTipoEvento() As String: codTipoEvento = This.codTipoEvento: End Property
Property Let codTipoEvento(pValue As String): This.codTipoEvento = pValue: End Property

Property Get codUsuarioErp() As String: codUsuarioErp = This.codUsuarioErp: End Property
Property Let codUsuarioErp(pValue As String): This.codUsuarioErp = pValue: End Property

Property Get nomeUsuarioErp() As String: nomeUsuarioErp = This.nomeUsuarioErp: End Property
Property Let nomeUsuarioErp(pValue As String): This.nomeUsuarioErp = pValue: End Property

Property Get caminhoDeProcessados() As String: caminhoDeProcessados = This.caminhoDeProcessados: End Property
Property Let caminhoDeProcessados(pValue As String): This.caminhoDeProcessados = pValue: End Property

Property Get strFileName() As String: strFileName = This.strFileName: End Property
Property Let strFileName(pValue As String): This.strFileName = pValue: End Property


Public Function criarERP()
Dim objERP As Collection: Set objERP = New Collection

    objERP.add "{"
    objERP.add "    ""Chave"": """ & This.ChvAcesso & ""","
    objERP.add "    ""Comando"": """ & This.Comando & """, "
    objERP.add "    ""data"": " & This.dhEmi & ","
    objERP.add "    ""codUsuarioErp"": """ & This.codUsuarioErp & """, "
    objERP.add "    ""nomeUsuarioErp"": """ & This.nomeUsuarioErp & """"
    objERP.add "}"

This.strFileName = This.ChvAcesso & ".erp.json"

execucao objERP, This.strFileName, strApp:="Notepad.exe", pOperacao:=opNone, strFilePath:=This.caminhoDeProcessados

End Function

Public Function criarManifesto()
Dim objManifesto As Collection: Set objManifesto = New Collection

    objManifesto.add "{"
    objManifesto.add "    ""chaveAcessoNFe"": """ & This.ChvAcesso & ""","
    objManifesto.add "    ""codTipoEvento"": """ & This.codTipoEvento & """, "
    objManifesto.add "    ""data"": " & This.dhEmi & ","
    objManifesto.add "    ""codUsuarioErp"": """ & This.codUsuarioErp & """, "
    objManifesto.add "    ""nomeUsuarioErp"": """ & This.nomeUsuarioErp & """"
    objManifesto.add "}"

This.strFileName = This.ChvAcesso & "-" & codTipoEvento & ".man.json"

execucao objManifesto, This.strFileName, strApp:="Notepad.exe", pOperacao:=opNone, strFilePath:=This.caminhoDeProcessados

End Function

Public Function NewEnum() As IUnknown
    Set NewEnum = pCol.[_NewEnum]
End Function

Private Sub Class_Initialize()
    Set pCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set pCol = Nothing
End Sub

Public Sub add(ByVal rec As clsConexaoNfeCteDados, Optional ByVal key As Variant, Optional ByVal before As Variant, Optional ByVal after As Variant)
    pCol.add rec, key, before, after
End Sub

Public Sub all(ByVal rec As clsConexaoNfeCteDados)
    pCol.add rec
End Sub

Public Function count() As Long
    count = pCol.count
End Function

Public Sub remove(ByVal i As Integer)
    pCol.remove i
End Sub

Public Function clsConexaoNfeCteDados(ByVal i As Variant) As clsConexaoNfeCteDados
    Set clsConexaoNfeCteDados = pCol.Item(i)
End Function

Public Property Get Itens() As Collection
    Set Itens = pCol
End Property

Public Property Get Item(i As Variant) As clsConexaoNfeCteDados
    Set Item = pCol(i)
End Property
