Attribute VB_Name = "01_modFernanda"
Option Compare Database

Private Const qrySelectRegistroValido As String = _
        "SELECT DISTINCT tblDadosConexaoNFeCTe.ChvAcesso, tblDadosConexaoNFeCTe.dhEmi FROM tblDadosConexaoNFeCTe WHERE (((Len([ChvAcesso]))>0) AND ((Len([dhEmi]))>0) AND ((tblDadosConexaoNFeCTe.registroValido)=1));"


''----------------------------
'' ### EXEMPLOS DE FUNÇÕES
''
'' exemploa_criacao_arquivos_json

''----------------------------


Sub exemplos_criacao_arquivos_json()
Dim s As New clsConexaoNfeCte

    '' NO PROCESSAMENTO DO ARQUIVO DE XML
    s.criar_ArquivoJson opFlagLancadaERP, qrySelectRegistroValido, "C:\temp\20210524\"
    
    '' SELEÇÃO PELO USUARIO
    s.criar_ArquivoJson opManifesto, qrySelectRegistroValido, "C:\temp\20210524\"
    
    
    MsgBox "Concluido!", vbOKOnly + vbInformation, "teste_arquivos_json"
    
Cleanup:

    Set s = Nothing

End Sub
