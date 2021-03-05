VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDadosConexaoNFeCTe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ClassType
    ID_Tipo As Integer
    codMod As Integer
    codIntegrado As Integer
    dhEmi As String
    CNPJ_emit As String
    Razao_emit As String
    CNPJ_Rem As String
    CPNJ_Dest As String
    CaminhoDoArquivo As String
    Chave As String
    Comando As String
End Type


Private This As ClassType

Property Get ID_Tipo() As Integer:                   ID_Tipo = This.ID_Tipo:          End Property
Property Let ID_Tipo(pValue As Integer):             This.ID_Tipo = pValue:              End Property
Property Get codMod() As Integer:                   codMod = This.codMod:          End Property
Property Let codMod(pValue As Integer):             This.codMod = pValue:              End Property
Property Get codIntegrado() As Integer:                   codIntegrado = This.codIntegrado:          End Property
Property Let codIntegrado(pValue As Integer):             This.codIntegrado = pValue:              End Property
Property Get dhEmi() As String:                   dhEmi = This.dhEmi:          End Property
Property Let dhEmi(pValue As String):             This.dhEmi = pValue:              End Property
Property Get CNPJ_emit() As String:                   CNPJ_emit = This.CNPJ_emit:          End Property
Property Let CNPJ_emit(pValue As String):             This.CNPJ_emit = pValue:              End Property
Property Get Razao_emit() As String:                   Razao_emit = This.Razao_emit:          End Property
Property Let Razao_emit(pValue As String):             This.Razao_emit = pValue:              End Property
Property Get CNPJ_Rem() As String:                   CNPJ_Rem = This.CNPJ_Rem:          End Property
Property Let CNPJ_Rem(pValue As String):             This.CNPJ_Rem = pValue:              End Property
Property Get CPNJ_Dest() As String:                   CPNJ_Dest = This.CPNJ_Dest:          End Property
Property Let CPNJ_Dest(pValue As String):             This.CPNJ_Dest = pValue:              End Property
Property Get CaminhoDoArquivo() As String:                   CaminhoDoArquivo = This.CaminhoDoArquivo:          End Property
Property Let CaminhoDoArquivo(pValue As String):             This.CaminhoDoArquivo = pValue:              End Property
Property Get Chave() As String:                   Chave = This.Chave:          End Property
Property Let Chave(pValue As String):             This.Chave = pValue:              End Property
Property Get Comando() As String:                   Comando = This.Comando:          End Property
Property Let Comando(pValue As String):             This.Comando = pValue:              End Property


Sub cadastroDados(pFolder As String)

Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
On Error GoTo cadastroDados_err

Dim fileName As Variant
Dim tmpFileName As String

    For Each fileName In GetFilesInSubFolders(pFolder)
        XDoc.Load fileName
        
        CaminhoDoArquivo = fileName
        
        Debug.Print "###### - " & fileName
        
        '' #SEM_IDENTIFICACAO
        If (XDoc.getElementsByTagName("ide")(0) Is Nothing) Then GoTo moverItemSemIdentificacao
        
        '' #SEM_REMETENTE
        If (XDoc.getElementsByTagName("rem")(0) Is Nothing) Then
            CNPJ_Rem = ""
        Else
            CNPJ_Rem = XDoc.getElementsByTagName("rem")(0).getElementsByTagName("CNPJ")(0).text
        End If
                
        '' #SEM_CHAVE
        If (XDoc.getElementsByTagName("infCTeNorm")(0) Is Nothing) Then
            Chave = ""
        Else
            Chave = XDoc.getElementsByTagName("infCTeNorm")(0).getElementsByTagName("infDoc")(0).getElementsByTagName("infNFe")(0).getElementsByTagName("chave")(0).text
        End If
        
        '' #SEM_CPNJ_Dest
        If (XDoc.getElementsByTagName("dest")(0).getElementsByTagName("CNPJ")(0) Is Nothing) Then
            If (XDoc.getElementsByTagName("dest")(0).getElementsByTagName("CPF")(0) Is Nothing) Then
                CPNJ_Dest = ""
            Else
                CPNJ_Dest = XDoc.getElementsByTagName("dest")(0).getElementsByTagName("CPF")(0).text
            End If
        Else
            CPNJ_Dest = XDoc.getElementsByTagName("dest")(0).getElementsByTagName("CNPJ")(0).text
        End If
                
'''############################################
'''############################################
'''############################################
                
'                cadastrar --->>>>>>>>>>> ''#Ailton - Implementação de método

'''############################################
'''############################################
'''############################################



'            .Parameters("ID_Tipo") = 6
'            .Parameters("codMod") = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("mod")(0).text
'            .Parameters("codIntegrado") = 0
'            .Parameters("dhEmi") = XDoc.getElementsByTagName("ide")(0).getElementsByTagName("dhEmi")(0).text
'            .Parameters("CNPJ_emit") = XDoc.getElementsByTagName("emit")(0).getElementsByTagName("CNPJ")(0).text
'            .Parameters("Razao_emit") = XDoc.getElementsByTagName("emit")(0).getElementsByTagName("xNome")(0).text
'            .Parameters("CNPJ_Rem") = CNPJ_Rem
'            .Parameters("CPNJ_Dest") = CPNJ_Dest
'            .Parameters("CaminhoDoArquivo") = fileName
'            .Parameters("Chave") = Chave

        
        If (Len(Chave) > 0) Then carregarERP
        
ContinueForLoop:
    Next fileName

cadastroDados_Fim:
    Set XDoc = Nothing
    Exit Sub
cadastroDados_err:
    MsgBox Err.description
    Resume cadastroDados_Fim

moverItemSemIdentificacao:
    Dim fDestino As String: fDestino = Replace(getPath(CaminhoDoArquivo), pegarValorDoParametro("caminhoDeColeta"), pegarValorDoParametro("caminhoDeProcessados"))
    CreateDir fDestino
    FileSystem.moveFile fileName, fDestino & getFileNameAndExt(CaminhoDoArquivo)
    GoTo ContinueForLoop



End Sub


Sub cadastrar(obj As clsDadosConexaoNFeCTe)

Dim db As DAO.Database: Set db = CurrentDb
Dim qdfDadosConexaoNFeCTe As DAO.QueryDef: Set qdfDadosConexaoNFeCTe = db.QueryDefs("INSERT_tblDadosConexaoNFeCTe")


    With qdfDadosConexaoNFeCTe
        .Parameters("ID_Tipo") = obj.ID_Tipo
        .Parameters("codMod") = obj.codMod
        .Parameters("codIntegrado") = obj.codIntegrado
        .Parameters("dhEmi") = obj.dhEmi
        .Parameters("CNPJ_emit") = obj.CNPJ_emit
        .Parameters("Razao_emit") = obj.Razao_emit
        .Parameters("CNPJ_Rem") = obj.CNPJ_Rem
        .Parameters("CPNJ_Dest") = obj.CPNJ_Dest
        .Parameters("CaminhoDoArquivo") = obj.CaminhoDoArquivo
        .Parameters("Chave") = obj.Chave
        
        ''#Ailton - Implementar dados do campo - comando
        .Parameters("Comando") = ""
        
        .Execute
    End With

End Sub


Sub carregarERP()
Dim c As New Collection: Dim s As New clsJsonErp

    s.strFilePath = getPath(CaminhoDoArquivo)
    
    s.Chave = Chave
    s.Comando = "ATUALIZAR"
    s.data = timestamp
    s.codUsuarioErp = pegarValorDoParametro("UsuarioErpCod")
    s.nomeUsuarioErp = pegarValorDoParametro("UsuarioErpNome")
    
    execucao s.criarERP, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath

End Sub

Sub carregarManifesto()
Dim c As New Collection: Dim s As New clsJsonManisfesto

    s.strFilePath = getPath(CaminhoDoArquivo)
    
    s.chaveAcessoNFe = Chave
    s.codTipoEvento = Comando
    s.data = timestamp
    s.codUsuarioErp = pegarValorDoParametro("UsuarioErpCod")
    s.nomeUsuarioErp = pegarValorDoParametro("UsuarioErpNome")
    
    execucao s.criarManifesto, s.strFileName, strApp:="Notepad.exe", pOperacao:=opNotepad, strFilePath:=s.strFilePath

End Sub


'Public Function Tipos(ByRef pCol As Collection)
'
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'1 - NF-e Importação')"
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'2 - NF-e Consumo')"
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'3 - NF-e com código Sisparts')"
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(55,'4 - NF-e Retorno Armazém')"
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'5 - NF-e')"
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(55,'6 - NF-e Transferência com código Sisparts')"
'    pCol.Add "INSERT INTO tblTipos (codMod,Descricao) VALUES(0,'7 - NF-e Transferência Uso/Consumo com código Sisparts')"
'
'End Function
