Attribute VB_Name = "azs_VALIDAR_DADOS"
Option Compare Database

Sub teste_cadastro()
Dim s As New clsProcessamentoDados

s.ProcessamentoTransferir "tblCompraNFItem"

End Sub


Sub teste_qryComprasCTe_Update_AjustesCampos()

'' BANCO_DESTINO
Dim strUsuarioNome As String: strUsuarioNome = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Usuario'")
Dim strUsuarioSenha As String: strUsuarioSenha = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Senha'")
Dim strOrigem As String: strOrigem = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Origem'")
Dim strBanco As String: strBanco = DLookup("[ValorDoParametro]", "[tblParametros]", "[TipoDeParametro]='BancoDados_Banco'")
Dim dbDestino As New Banco

'' BANCO_LOCAL
Dim Scripts As New clsConexaoNfeCte
Dim qryComprasCTe_Update_AjustesCampos As String: qryComprasCTe_Update_AjustesCampos = "UPDATE tblCompraNF SET tblCompraNF.HoraEntd_CompraNF = NULL ,tblCompraNF.IDVD_CompraNF = NULL WHERE (((tblCompraNF.ChvAcesso_CompraNF) IN (pLista_ChvAcesso_CompraNF)));"


    '' BANCO_DESTINO
    dbDestino.Start strUsuarioNome, strUsuarioSenha, strOrigem, strBanco, drSqlServer
    
    dbDestino.SqlExecute Replace(qryComprasCTe_Update_AjustesCampos, "pLista_ChvAcesso_CompraNF", carregarComprasCTe)
    

End Sub


Sub teste01()
Dim s As New clsProcessamentoDados

Dim pCaminho As String: pCaminho = _
    "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"

s.ProcessamentoDeArquivo pCaminho, opDadosGerais

End Sub


Sub TestStub()
          
Dim pCaminho As String: pCaminho = _
    "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"
          
          
    Dim objXML As Object, node As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    objXML.async = False: objXML.validateOnParse = False

    If Not objXML.Load(pCaminho) Then  'strXML is the string with XML'
        Err.Raise objXML.parseError.errorCode, , objXML.parseError.reason

    Else
        Set node = objXML.selectSingleNode("cteProc")
        Stop

    End If
End Sub


Sub azsLerNodes()

Dim pCaminho As String: pCaminho = _
    "C:\xmls\68.365.5010003-77 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\recebimento\42210348740351012767570000021186731952977908-cteproc.xml"

Dim objXML As MSXML2.DOMDocument60: Set objXML = New MSXML2.DOMDocument60
    
'Dim XDoc As Object: Set XDoc = CreateObject("MSXML2.DOMDocument"): XDoc.async = False: XDoc.validateOnParse = False
objXML.async = False: objXML.validateOnParse = False
objXML.Load pCaminho

'Dim Nodes As IXMLDOMNodeList: Set Nodes = objXML.childNodes

Dim objNode As IXMLDOMNode

   For Each objNode In objXML.childNodes
      If objNode.NodeType = NODE_TEXT Then
          If objNode.ParentNode.nodeName = "pICMS" Then
            Debug.Print CStr(objNode.NodeValue)
          End If
      End If
   Next objNode
End Sub


Private Sub gerar_ArquivosDeValidacaoDeCampos()
Dim db As DAO.Database: Set db = CurrentDb
Dim rstRegistros As DAO.Recordset
Dim rstItens As DAO.Recordset
Dim sqlRegistros As String: sqlRegistros = "Select * from tblCompraNF where ChvAcesso_CompraNF = "
Dim sqlItens As String: sqlItens = "Select * from tblCompraNFItem where ChvAcesso_CompraNF = "
Dim arquivos As New Collection

Dim item As Variant
Dim TMP As String
''42210300634453001303570010001139451001171544
 arquivos.add "42210300634453001303570010001139451001171544" '' 57

'arquivos.add "32210368365501000296550000000639051001364146"

'arquivos.add "32210304884082000569570000040073831040073834"
'arquivos.add "42210220147617000494570010009539201999046070"
'arquivos.add "32210368365501000296550000000638811001361356"
'arquivos.add "42210212680452000302550020000886301507884230"

'arquivos.Add "32210368365501000296550000000638841001361501"


For Each item In arquivos

    
    TMP = sqlRegistros & "'" & CStr(item) & "'"
    Set rstRegistros = db.OpenRecordset(TMP)
    
    Do While Not rstRegistros.EOF
        
        TMP = ""
        For i = 0 To rstRegistros.Fields.count - 1
            TMP = rstRegistros.Fields(i).Name & vbTab & rstRegistros.Fields(i).value
            TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", TMP
        Next i

        TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", vbNewLine & "#############################" & vbNewLine

        TMP = ""
        TMP = sqlItens & "'" & CStr(item) & "'"
        Debug.Print TMP
        
        Set rstItens = db.OpenRecordset(TMP)
        Do While Not rstItens.EOF
            For i = 0 To rstItens.Fields.count - 1
                TMP = rstItens.Fields(i).Name & vbTab & rstItens.Fields(i).value
                TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", TMP
            Next i
            
            TextFile_Append CurrentProject.path & "\" & CStr(item) & ".txt", vbNewLine & "#############################" & vbNewLine
            
            rstItens.MoveNext
            DoEvents
        Loop

        Debug.Print "Concluido! - " & CStr(item) & ".txt"
        rstRegistros.MoveNext
        DoEvents
        TMP = ""
    Loop
    
    rstRegistros.Close
    rstItens.Close
Next

Debug.Print "Concluido!"
Set rstRegistros = Nothing
Set rstItens = Nothing

End Sub

'Private Sub criarConsultasParaTestes()
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rstOrigem As DAO.Recordset
'Dim strSQL As String
'Dim qrySelectTabelas As String: qrySelectTabelas = "Select Distinct tabela from tblOrigemDestino order by tabela"
'Dim tabela As Variant
'
''' CRIAR CONSULTA PARA VALIDAR DADOS PROCESSADOS
'For Each tabela In carregarParametros(qrySelectTabelas)
'    strSQL = "Select "
'    Set rstOrigem = db.OpenRecordset("Select distinct Destino from tblOrigemDestino where tabela = '" & tabela & "'")
'    Do While Not rstOrigem.EOF
'        strSQL = strSQL & strSplit(rstOrigem.Fields("Destino").value, ".", 1) & ","
'        rstOrigem.MoveNext
'    Loop
'
'    strSQL = left(strSQL, Len(strSQL) - 1) & " from " & tabela
'    qryDeleteExists "qry_" & tabela
'    qryCreate "qry_" & tabela, strSQL
'Next tabela
'
'db.Close: Set db = Nothing
'
'End Sub
