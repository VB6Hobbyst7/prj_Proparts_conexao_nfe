VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJsonManisfesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private Type ClassType
    strFilePath As String
    strFileName As String
    chaveAcessoNFe As String
    codTipoEvento As String
    data As String
    codUsuarioErp As String
    nomeUsuarioErp As String
    
End Type

Private This As ClassType

Property Get strFilePath() As String:                   strFilePath = This.strFilePath:          End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:              End Property

Property Get strFileName() As String:                   strFileName = This.strFileName:          End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:              End Property

Property Get chaveAcessoNFe() As String:                   chaveAcessoNFe = This.chaveAcessoNFe:          End Property
Property Let chaveAcessoNFe(pValue As String):             This.chaveAcessoNFe = pValue:              End Property

Property Get codTipoEvento() As String:                   codTipoEvento = This.codTipoEvento:          End Property
Property Let codTipoEvento(pValue As String):             This.codTipoEvento = pValue:              End Property

Property Get data() As String:                   data = This.data:          End Property
Property Let data(pValue As String):             This.data = pValue:              End Property

Property Get codUsuarioErp() As String:                   codUsuarioErp = This.codUsuarioErp:          End Property
Property Let codUsuarioErp(pValue As String):             This.codUsuarioErp = pValue:              End Property

Property Get nomeUsuarioErp() As String:                   nomeUsuarioErp = This.nomeUsuarioErp:          End Property
Property Let nomeUsuarioErp(pValue As String):             This.nomeUsuarioErp = pValue:              End Property


Public Function criarManifesto() 'ByRef pCol As Collection
Set criarManifesto = New Collection

criarManifesto.add "{"
criarManifesto.add "    ""chaveAcessoNFe"": """ & chaveAcessoNFe & ""","
criarManifesto.add "    ""codTipoEvento"": """ & codTipoEvento & """, "
criarManifesto.add "    ""data"": " & data & ","
criarManifesto.add "    ""codUsuarioErp"": """ & codUsuarioErp & """, "
criarManifesto.add "    ""nomeUsuarioErp"": """ & nomeUsuarioErp & """"
criarManifesto.add "}"

This.strFileName = chaveAcessoNFe & "-" & codTipoEvento & ".man.json"

End Function
