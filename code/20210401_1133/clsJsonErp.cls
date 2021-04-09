VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsJsonErp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Option Explicit

Private Type ClassType
    strFilePath As String
    strFileName As String
    Chave As String
    Comando As String
    data As String
    codUsuarioErp As String
    nomeUsuarioErp As String
    
End Type


Private This As ClassType

Property Get strFilePath() As String:                   strFilePath = This.strFilePath:          End Property
Property Let strFilePath(pValue As String):             This.strFilePath = pValue:              End Property

Property Get strFileName() As String:                   strFileName = This.strFileName:          End Property
Property Let strFileName(pValue As String):             This.strFileName = pValue:              End Property

Property Get Chave() As String:                   Chave = This.Chave:          End Property
Property Let Chave(pValue As String):             This.Chave = pValue:              End Property

Property Get Comando() As String:                   Comando = This.Comando:          End Property
Property Let Comando(pValue As String):             This.Comando = pValue:              End Property

Property Get data() As String:                   data = This.data:          End Property
Property Let data(pValue As String):             This.data = pValue:              End Property

Property Get codUsuarioErp() As String:                   codUsuarioErp = This.codUsuarioErp:          End Property
Property Let codUsuarioErp(pValue As String):             This.codUsuarioErp = pValue:              End Property

Property Get nomeUsuarioErp() As String:                   nomeUsuarioErp = This.nomeUsuarioErp:          End Property
Property Let nomeUsuarioErp(pValue As String):             This.nomeUsuarioErp = pValue:              End Property


Public Function criarERP()
Set criarERP = New Collection

criarERP.add "{"
criarERP.add "    ""Chave"": """ & Chave & ""","
criarERP.add "    ""Comando"": """ & Comando & """, "
criarERP.add "    ""data"": " & data & ","
criarERP.add "    ""codUsuarioErp"": """ & codUsuarioErp & """, "
criarERP.add "    ""nomeUsuarioErp"": """ & nomeUsuarioErp & """"
criarERP.add "}"

This.strFileName = Chave & ".erp.json"

End Function

