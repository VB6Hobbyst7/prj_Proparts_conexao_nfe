<<<<<<< HEAD:referencia/code/modFuncoesGerais.bas
=======
<<<<<<< HEAD
>>>>>>> ca95ee3e8bcb0745be1525054e4155ff5a288f06:referencia/modFuncoesGerais.bas
Attribute VB_Name = "modFuncoesGerais"
Option Compare Database
Global Var_Acesso As Integer
Global txtEnvioBol As Integer
Global frmNomeForm As String
Global IDVDAtu As Long
'############################################
'######### INICIO DAS APIs DO WINDOWS #######
'############################################
'######## ==>>> API DE SELE��O DE ARQUIVOS! #####################################
'######## ==>>>>>>>> Dependente da DLL COMDLG32.DLL - REQUER W95 OU SUPERIOR !


'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Type OPENFILENAME
'    lStructSize As Long
'    hwndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    nFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    Flags As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

#If VBA7 Then
    Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As LongPtr
    lpTemplateName As String
    End Type
#Else
    Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
#End If

'###############################################################################

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260

'Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
'Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'
'Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
'        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
'        As String, ByVal lpOutPath As String) As Long

#If VBA7 Then
    Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Public Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Public Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Public Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    
    Public Declare PtrSafe Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
            "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
            As String, ByVal lpOutPath As String) As Long

#Else
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    
    Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
            "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
            As String, ByVal lpOutPath As String) As Long
#End If


'############################################
'######### FINAL DAS  APIs DO WINDOWS #######
'############################################
Public Type T_MySql
 campo As String
 valor As Variant
End Type

Public Enum T_Valor
 Num�rico = 0
 texto = 1
 Auto_Detec��o = 2
End Enum

Dim MySqlConn As ADODB.Connection
Dim LCampos As String
Dim Campos() As T_MySql
Global ulibera As String

'Permiss�es
Global txtPer
Option Explicit

'*****************************************************
#If VBA7 Then
Const ChaveReg = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\Theme"
#Else
Const ChaveReg = "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Common\Theme"
#End If
 
Enum eCor
   Azul = 1
   Prata = 2
   Preto = 3
End Enum
'*****************************************************


Function EstaAbertoForm(NomeFormul�rio As String) As Integer
On Error GoTo Err_Est�Aberto
Dim N�meroDeFormul�rios
Dim x As Integer
    N�meroDeFormul�rios = Forms.count                                            ' formul�rios.
    For x = 0 To N�meroDeFormul�rios - 1
        If Forms(x).Name = NomeFormul�rio Then
            EstaAbertoForm = -1
            Exit Function
        Else
            EstaAbertoForm = 0
        End If
    Next x
Exit_Est�Aberto:
    Exit Function
Err_Est�Aberto:
    MsgBox Error$
    Resume Exit_Est�Aberto
End Function

Public Sub fncCorAccess(cor As eCor)
Dim ChaveReg As String
Dim objReg As Object: Set objReg = CreateObject("wscript.shell")
'Gravar valor na chave THEME
objReg.RegWrite ChaveReg, cor, "REG_DWORD"
Set objReg = Nothing
End Sub
'Para alterar a cor , basta executar o procedimento: Call fncCorAccess(Preto)
'Reiniciando o Access
'O �nico por�m � que a cor s� � alterada na abertura seguinte do Access.  Para criar um efeito imediato na configura��o, montei um c�digo que reinicia automaticamente a aplica��o. Aqui est�:

Public Sub fncReiniciandoAplicativo()
Dim strLocal As String
Dim objWs As Object
'------------------------
'Reiniciando
'------------------------
Set objWs = CreateObject("wscript.shell")
strLocal = CurrentProject.path & "\" & CurrentProject.Name
strLocal = Chr(34) & "MSACCESS.EXE" & Chr(34) & " " & Chr(34) _
           & strLocal & Chr(34)
'0 - oculto / 5 - vis�vel
objWs.Run strLocal, 5, "false"
'-------------------------------------
'Fecha atual
'-------------------------------------
Application.Quit acQuitSaveAll
End Sub


Function AbrirCombo()
On Error GoTo Err_AbrirCombo
    SendKeysNovo ("{F4}")
Exit_AbrirCombo:
    Exit Function
Err_AbrirCombo:
    MsgBox Error$
    Resume Exit_AbrirCombo
End Function

Function AbrirFormul�rio(NomeFormul�rio As String, Crit�rio As String) As Integer
On Error GoTo Err_AbrirFormul�rio
    DoCmd.OpenForm NomeFormul�rio, , , Crit�rio
Exit_AbrirFormul�rio:
    Exit Function
Err_AbrirFormul�rio:
    MsgBox Error$
    Resume Exit_AbrirFormul�rio
End Function

Function AbrirObj(varTPObj As Variant, varNmObj As Variant) As Integer
On Error GoTo Err_AbrirObj
    Select Case varTPObj
    Case "Formul�rio"
        DoCmd.OpenForm varNmObj
    Case "Relat�rio"
        DoCmd.OpenReport varNmObj, A_PREVIEW
    End Select
Exit_AbrirObj:
    Exit Function
Err_AbrirObj:
    MsgBox Error
    Resume Exit_AbrirObj
End Function

Function Acesso() As Integer
On Error GoTo Err_Acesso
Dim db  As Database, rs As Recordset, x As Long, Doc As Document
Dim y
Dim i
  
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset("tblAM_Acesso")
    Set Doc = db.Containers(7).Documents("tblAM_Acesso")
    x = 0
    While Not rs.EOF
        x = x + 1
        rs.MoveNext
    Wend
    If x < 30 Then
        y = MsgBox("Esta � a " & (x + 1) & "� vez que voc� acessa a demonstra��o deste produto!", 64, "Aten��o")
        y = AbrirFormul�rio("frmIni", "")
        rs.AddNew
        rs!DT_Acesso = Date
        rs.Update
        Var_Acesso = 1
    If Doc.UserName = "SUPERVISOR" Then
      If MsgBox("Deseja liberar o acesso?", 36, "Libera��o do Sistema") = 6 Then
        For i = x To 31
          rs.AddNew
          rs!DT_Acesso = Date
          rs.Update
        Next i
      End If
    End If
    Else
        If x = 30 Then
            y = MsgBox("Esta � a " & (x + 1) & "� vez que voc� acessa a demonstra��o deste produto!", 48, "Aten��o")
            If Doc.UserName = "SUPERVISOR" Then
              MsgBox Trim(Mid(str$(CLng(Date)), 1, 3) & Chr(Day(Date) / 2 + 65) & Mid(str$(CLng(Date)), 4, 4))
            End If
            Var_Acesso = Senha()
        Else
            Var_Acesso = 1
        End If
    End If
    rs.Close
    Acesso = Var_Acesso
Exit_Acesso:
    Exit Function
Err_Acesso:
    MsgBox Error$
    Resume Exit_Acesso
End Function

Function Ajuda()
On Error GoTo Err_Ajuda
    SendKeysNovo ("{F1}")
Exit_Ajuda:
    Exit Function
Err_Ajuda:
    MsgBox Error$
    Resume Exit_Ajuda
End Function

Function AplicarFiltro(NomeFiltro As String) As Integer
'Aplica um filtro salvo como consulta
'Domingos Taffarello - 19/10/96
On Error GoTo Err_AplicarFiltro
    DoCmd.ApplyFilter NomeFiltro
Exit_AplicarFiltro:
    Exit Function
Err_AplicarFiltro:
    MsgBox Error$
    Resume Exit_AplicarFiltro
End Function

Function AvisarSa�da()
On Error GoTo Err_AvisarSa�da
Dim x As Integer
    If MsgBox("Deseja sair deste sistema?", 276, "Aviso de sa�da") = 6 Then
        x = Sair()
    End If
Exit_AvisarSa�da:
    Exit Function
Err_AvisarSa�da:
    MsgBox Error$
    Resume Exit_AvisarSa�da
End Function

Function ChecarAcesso()
'If (Var_Acesso = 0 And DCount("ID_Acesso", "tblAM_Acesso") >= 30) Or IsNull(Var_Acesso) Then X = Sair()
End Function

Function CompletaString(varTexto As String, strCompleta As String, intLargura As Integer, strLado As String) As String
'Completa a string varTexto com o caracter strCompleta, o n�mero de vezes at�
'atingir o comprimento intLargura. strLado indica se ser� � esquerda ou � direita
'Domingos Taffarello - 31/08/96
On Error GoTo Err_CompletaString
Dim strTemp As String
    strTemp = Trim((varTexto))
    While Len(strTemp) < intLargura
        If strLado = "E" Then
            strTemp = strCompleta & strTemp
        Else
            strTemp = strTemp & strCompleta
        End If
    Wend
    CompletaString = strTemp
Exit_CompletaString:
    Exit Function
Err_CompletaString:
    MsgBox Error$
    Resume Exit_CompletaString
End Function

Function EAN13(x) As String
On Error GoTo Err_EAN13
Dim p
Dim i
     If Len(Trim$(x)) < 12 Then
        MsgBox "Foi digitado um n�mero com menos de 12 d�gitos!"
        EAN13 = x
        Exit Function
     End If
     If Len(Trim$(x)) > 12 Then
        x = Mid(Trim$(x), 1, 12)
'        MsgBox "Foi digitado um n�mero com mais de 12 d�gitos!"
'        EAN13 = X
'        Exit Function
    End If
    p = 0
    For i = 1 To 12
        p = p + Val(Mid(Trim$(x), i, 1) * IIf(i = 1, 1, IIf((i Mod 2) = 0, 3, 1)))
    Next i
    EAN13 = x & Trim(str(((p \ 10 + IIf((p Mod 10) > 0, 1, 0)) * 10) - p))
Exit_EAN13:
    Exit Function
Err_EAN13:
    MsgBox Error$
    Resume Exit_EAN13
End Function

Function Excluir()
On Error GoTo Err_Excluir
    SendKeysNovo ("^{-}")
Exit_Excluir:
    Exit Function
Err_Excluir:
    MsgBox Error$
    Resume Exit_Excluir
End Function

Function ExibeSenha()
    MsgBox Trim(Mid(str$(CLng(Date)), 1, 3) & Chr(Day(Date) / 2 + 65) & Mid(str$(CLng(Date)), 4, 4))
End Function

Function Extenso(nValor)
'* Extenso()
'* Sintaxe..: Extenso(nValor) -> cExtenso
'* Descri��o: Retorna uma s�rie de caracteres contendo a forma extensa
'*            do valor passado como argumento.
'* Autoria..: Eng. Cesar Costa e Dalicio Guiguer Filho
'* Linguagem: Access Basic
'* Data.....: Fevereiro/1994
On Error GoTo Err_Extenso
'Faz a valida��o do argumento
  If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
    Exit Function
  End If
'Declara as vari�veis da fun��o
  Dim nContador, nTamanho As Integer
  Dim cValor, cParte, cFinal As String
  ReDim aGrupo(4), aTexto(4) As String
'Define matrizes com extensos parciais
  ReDim aUnid(19) As String
  aUnid(1) = "UM ": aUnid(2) = "DOIS ": aUnid(3) = "TRES "
  aUnid(4) = "QUATRO ": aUnid(5) = "CINCO ": aUnid(6) = "SEIS "
  aUnid(7) = "SETE ": aUnid(8) = "OITO ": aUnid(9) = "NOVE "
  aUnid(10) = "DEZ ": aUnid(11) = "ONZE ": aUnid(12) = "DOZE "
  aUnid(13) = "TREZE ": aUnid(14) = "QUATORZE ": aUnid(15) = "QUINZE "
  aUnid(16) = "DEZESSEIS ": aUnid(17) = "DEZESSETE ": aUnid(18) = "DEZOITO "
  aUnid(19) = "DEZENOVE "
  ReDim aDezena(9) As String
  aDezena(1) = "DEZ ": aDezena(2) = "VINTE ": aDezena(3) = "TRINTA "
  aDezena(4) = "QUARENTA ": aDezena(5) = "CINQUENTA "
  aDezena(6) = "SESSENTA ": aDezena(7) = "SETENTA ": aDezena(8) = "OITENTA "
  aDezena(9) = "NOVENTA "
  ReDim aCentena(9) As String
  aCentena(1) = "CENTO ":  aCentena(2) = "DUZENTOS "
  aCentena(3) = "TREZENTOS ": aCentena(4) = "QUATROCENTOS "
  aCentena(5) = "QUINHENTOS ": aCentena(6) = "SEISCENTOS "
  aCentena(7) = "SETECENTOS ": aCentena(8) = "OITOCENTOS "
  aCentena(9) = "NOVECENTOS "
'Divide o valor em v�rios grupos
  cValor = Format$(nValor, "0000000000.00")
  aGrupo(1) = Mid$(cValor, 2, 3)
  aGrupo(2) = Mid$(cValor, 5, 3)
  aGrupo(3) = Mid$(cValor, 8, 3)
  aGrupo(4) = "0" + Mid$(cValor, 12, 2)
'Processa cada grupo
  For nContador = 1 To 4
    cParte = aGrupo(nContador)
    nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
    If nTamanho = 3 Then
      If right$(cParte, 2) <> "00" Then
        aTexto(nContador) = aTexto(nContador) + aCentena(left(cParte, 1)) + "E "
        nTamanho = 2
      Else
        aTexto(nContador) = aTexto(nContador) + IIf(left$(cParte, 1) = "1", "CEM ", aCentena(left(cParte, 1)))
      End If
    End If
    If nTamanho = 2 Then
      If Val(right(cParte, 2)) < 20 Then
        aTexto(nContador) = aTexto(nContador) + aUnid(right(cParte, 2))
      Else
        aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
        If right$(cParte, 1) <> "0" Then
          aTexto(nContador) = aTexto(nContador) + "E "
          nTamanho = 1
        End If
      End If
    End If
    If nTamanho = 1 Then
      aTexto(nContador) = aTexto(nContador) + aUnid(right(cParte, 1))
    End If
  Next
'Gera o formato final do texto
  If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
    cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS")
  Else
    cFinal = ""
    cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "MILH�ES ", "MILH�O "), "")
    If Val(aGrupo(2) + aGrupo(3)) = 0 Then
      cFinal = cFinal + "DE "
    Else
      cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "MIL ", "")
    End If
    cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "REAL ", "REAIS ")
    cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS"), "")
  End If
  Extenso = cFinal
Exit_Extenso:
    Exit Function
Err_Extenso:
    MsgBox Error$
    Resume Exit_Extenso
End Function

'Function Fechar()
'On Error GoTo Err_Close
'    DoCmd.Close
'Exit_Close:
'    Exit Function
'Err_Close:
'    MsgBox Error$
'    Resume Exit_Close
'End Function

Function fncFechar(NomeFormul�rio As String)
On Error GoTo Err_FecharFormul�rio
    DoCmd.Close A_FORM, NomeFormul�rio
    'DoCmd.Close
Exit_FecharFormul�rio:
    Exit Function
Err_FecharFormul�rio:
    MsgBox Error$
    Resume Exit_FecharFormul�rio
End Function

Function Fix2(x As Variant) As Currency '' #AILTON - ARREDONDAMENTO
On Error GoTo Err_Fix2
    Fix2 = Int(Nz(x) * 100 + 0.5) / 100
Exit_Fix2:
    Exit Function
Err_Fix2:
    MsgBox Error$
    Resume Exit_Fix2
End Function
Function Fix4(x As Variant) As Currency
On Error GoTo Err_Fix4
    Fix4 = Int(Nz(x) * 10000 + 0.5) / 10000
Exit_Fix4:
    Exit Function
Err_Fix4:
    MsgBox Error$
    Resume Exit_Fix4
End Function


Function ForaDaLista()
On Error GoTo Err_ForaDaLista
Dim Response
    Response = DATA_ERRCONTINUE
    Screen.ActiveControl = Null
    SendKeysNovo ("{ESC}")
Exit_ForaDaLista:
    Exit Function
Err_ForaDaLista:
    MsgBox Error$
    Resume Exit_ForaDaLista
End Function

Function ImprimirRelat�rio(NomeRelat�rio As String, Crit�rio As String) As Integer
On Error GoTo Err_ImprimeNota_Click
    DoCmd.OpenReport NomeRelat�rio, A_PREVIEW, , Crit�rio
Exit_ImprimeNota_Click:
    Exit Function

Err_ImprimeNota_Click:
    MsgBox Error$
    Resume Exit_ImprimeNota_Click
End Function

Function Incluir(NomeForm As String, NomeControle As String)
On Error GoTo Err_Incluir
    DoCmd.OpenForm NomeForm
    Forms(NomeForm)(NomeControle).SetFocus
    DoCmd.GoToRecord , , A_NEWREC
Exit_Incluir:
    Exit Function
Err_Incluir:
    MsgBox Error$
    Resume Exit_Incluir
End Function

Function Inicializar()
'Fun��o de Inicializa��o
'Domingos Taffarello - 02/10/96
On Error GoTo Err_Inicializar

Dim x As Integer

DoCmd.OpenForm "frmLogin"
    
Exit_Inicializar:
    Exit Function
Err_Inicializar:
    MsgBox Error$
    Resume Next
    Resume Exit_Inicializar
End Function
Function IrParaControle(NomeControle As control) As Integer
On Error GoTo Err_IrParaControle
    NomeControle.SetFocus
Exit_IrParaControle:
    Exit Function
Err_IrParaControle:
    MsgBox Error$
    Resume Exit_IrParaControle
End Function

Function NullToZero(x As Variant) As Variant
On Error GoTo Err_NullToZero
    If IsNull(x) Then
        NullToZero = 0
    Else
        NullToZero = x
    End If
Exit_NullToZero:
    Exit Function
Err_NullToZero:
    MsgBox Error$
    Resume Exit_NullToZero
End Function

Function PLayWave(NomeArq As String) As Integer
On Error GoTo Err_PlayWave

    PLayWave = SndPlaySound(NomeArq, 1)
Exit_PlayWave:
    Exit Function
Err_PlayWave:
    MsgBox Error$
    Resume Exit_PlayWave
End Function

Function procurar(NomeForm As String, FindWhat As control, FindWhere As String, Find_A As Integer) As Integer
On Error GoTo Err_Procurar
Dim Invis�vel
    'DoCmd.DoMenuItem A_FORMBAR, A_FILE, A_SAVERECORD, , acMenuVer1X
    If IsNull(FindWhat) Then Exit Function
    DoCmd.OpenForm NomeForm
    'DoCmd.ShowAllRecords
    If Forms(NomeForm)(FindWhere).Visible = False Then
        Invis�vel = True
        Application.Echo False
        Forms(NomeForm)(FindWhere).Visible = True
    End If
        Forms(NomeForm)(FindWhere).SetFocus
        DoCmd.FindRecord FindWhat, Find_A
    If Invis�vel Then
        Application.Echo True
        SendKeysNovo ("{TAB}")
        Forms(NomeForm)(FindWhere).Visible = False
    End If
Exit_Procurar:
    Exit Function
Err_Procurar:
    MsgBox Error$
    Resume Exit_Procurar
End Function

Function Reconsultar()
On Error GoTo Err_Reconsultar
    Screen.ActiveForm.Requery
Exit_Reconsultar:
    Exit Function
Err_Reconsultar:
    MsgBox Error$
    Resume Exit_Reconsultar
End Function

Function ReconsultarControle(NomeControle As control) As Integer
On Error GoTo Err_ReconsultarControle
    NomeControle.Requery
Exit_ReconsultarControle:
    Exit Function
Err_ReconsultarControle:
    MsgBox Error$
    Resume Exit_ReconsultarControle
End Function
Function Sair()
On Error GoTo Err_Sair
    Application.Quit
Exit_Sair:
    Exit Function
Err_Sair:
    MsgBox Error$
    Resume Exit_Sair
End Function

Function Salvar()
On Error GoTo Err_Salvar
     DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_Salvar:
    Exit Function
Err_Salvar:
    MsgBox Error$
    Resume Exit_Salvar
End Function

Function Senha()
On Error GoTo Err_Senha
Dim rs As Recordset, db As Database
Dim z, y
    z = Trim(Mid(str$(CLng(Date)), 1, 3) & Chr(Day(Date) / 2 + 65) & Mid(str$(CLng(Date)), 4, 4))
    y = InputBox("Digite a senha:")
    
    If y = z Then
        Senha = 1
        Set db = DBEngine.Workspaces(0).Databases(0)
        Set rs = db.OpenRecordset("tblAM_Acesso")
        rs.AddNew
        rs!DT_Acesso = Date
        rs.Update
        rs.Close
    Else
        Senha = 0
    End If
Exit_Senha:
    Exit Function
Err_Senha:
    MsgBox Error$
    Resume Exit_Senha
End Function

Function TBOff(TB As String) As Integer
On Error GoTo Err_TBOff
    DoCmd.ShowToolbar TB, A_TOOLBAR_NO
Exit_TBOff:
    Exit Function
Err_TBOff:
    MsgBox Error$
    Resume Exit_TBOff
End Function

Function TBOn(TB As String) As Integer
On Error GoTo Err_TBOn
    DoCmd.ShowToolbar TB, A_TOOLBAR_YES
Exit_TBOn:
    Exit Function
Err_TBOn:
    MsgBox Error$
    Resume Exit_TBOn
End Function

Function Zoom()
On Error GoTo Err_Zoom
    SendKeysNovo ("+{F2}")
Exit_Zoom:
    Exit Function
Err_Zoom:
    MsgBox Error$
    Resume Exit_Zoom
End Function

Function DVCGC(CGC As String)
On Error GoTo Err_CGC
Dim intSoma, intSoma1, intSoma2, intInteiro As Long
Dim intNumero, intMais, i, intResto As Integer
Dim intDig1, intDig2 As Integer
Dim strCampo, strCaracter, strConf, strCGC As String
Dim dblDivisao As Double
intSoma = 0
intSoma1 = 0
intSoma2 = 0
intNumero = 0
intMais = 0
'Separa os d�gitos do CGC que ser�o multiplicados de 2 a 9.
'Retira a "/" da m�scara de entrada.
strCGC = right(CGC, 6)
strCGC = left(strCGC, 4)
strCampo = left(CGC, 8)
strCampo = right(strCampo, 4) & strCGC
For i = 2 To 9
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma1 = intSoma1 + intMais
Next i
'Separa os 4 primeiros d�gitos do CGC
strCampo = left(CGC, 4)
For i = 2 To 5
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma2 = intSoma2 + intMais
Next i
intSoma = intSoma1 + intSoma2
dblDivisao = intSoma / 11
intInteiro = Int(dblDivisao) * 11
intResto = intSoma - intInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If
intSoma = 0
intSoma1 = 0
intSoma2 = 0
intNumero = 0
intMais = 0
strCGC = right(CGC, 6)
strCGC = left(strCGC, 4)
strCampo = left(CGC, 8)
strCampo = right(strCampo, 3) & strCGC & intDig1
For i = 2 To 9
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma1 = intSoma1 + intMais
Next i
strCampo = left(CGC, 5)
For i = 2 To 6
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma2 = intSoma2 + intMais
Next i
intSoma = intSoma1 + intSoma2
dblDivisao = intSoma / 11
intInteiro = Int(dblDivisao) * 11
intResto = intSoma - intInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CGC esteja errado dispara a mensagem
If strConf <> right(CGC, 2) Then
    MsgBox "O d�gito do CNPJ n�o est� correto.", 16, "Aten��o"
    DVCGC = False
    Exit Function
End If
    DVCGC = True
Exit Function
Exit_CGC:
    Exit Function
Err_CGC:
    MsgBox Error$
    Resume Exit_CGC
End Function

Function DVCPF(CPF As String)
On Error GoTo Err_CPF
Dim lngSoma, lngInteiro As Long
Dim intNumero, intMais, i, intResto As Integer
Dim intDig1, intDig2 As Integer
Dim strCampo, strCaracter, strConf As String
Dim dblDivisao As Double
lngSoma = 0
intNumero = 0
intMais = 0
strCampo = left(CPF, 9)
'Inicia c�lculos do 1� d�gito
'A fun��o Right() separa os caracteres da direita
'A fun��o Left() separa os caracteres da esquerda
'A fun��o Int() retorna o valor inteiro de um campo num�rico
For i = 2 To 10
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If
strCampo = strCampo & intDig1
lngSoma = 0
intNumero = 0
intMais = 0
For i = 2 To 11
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> right(CPF, 2) Then
    MsgBox "O d�gito do CPF n�o est� correto.", 16, "Aten��o"
    DVCPF = False
    Exit Function
End If

DVCPF = True
Exit Function
Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function

Function CalculaData(data As Date) As Date
Dim sqlString As String
Dim rsFer As New ADODB.Recordset

Dim i As Integer

AbrirConexao

If Weekday(data) = 1 Then
  data = data + 1
End If
If Weekday(data) = 7 Then
  data = data + 2
End If

i = 0

For i = 0 To 7 Step 1
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(data, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    If rsFer.RecordCount = 0 Then
        rsFer.Close
        Exit For
    End If
    If rsFer.RecordCount <> 0 Then
    data = data + 1
      If Weekday(data) = 1 Then
        data = data + 1
      End If
      If Weekday(data) = 7 Then
        data = data + 2
      End If
    End If
    rsFer.Close
Next

CalculaData = data

End Function
Function CalculaDataCob(data As Date) As Date
Dim sqlString As String
Dim rsFer As New ADODB.Recordset
  
AbrirConexao

Dim i As Integer
'Dim db As DataBase
'Dim rs As Recordset
'Dim qry_Fer As QueryDef

If Weekday(data) = 1 Then
  data = data + 1
End If
If Weekday(data) = 7 Then
  data = data + 2
End If

i = 0

For i = 0 To 7 Step 1
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(data, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    'Set db = CurrentDb()
    'Set qry_Fer = db.QueryDefs("qryLctoFin_Feriados")
    'qry_Fer.Parameters(0) = Data
    'Set rs = qry_Fer.OpenRecordset()
    If rsFer.RecordCount = 0 Then
        rsFer.Close
        Exit For
    End If
    If rsFer.RecordCount <> 0 Then
    data = data + 1
      If Weekday(data) = 1 Then
        data = data + 1
      End If
      If Weekday(data) = 7 Then
        data = data + 2
      End If
    End If
    rsFer.Close
Next


'If Weekday(Data) + 1 = 7 Then
'    Data = Data + 3
'ElseIf Weekday(Data) + 1 = 1 Then
'    Data = Data + 3
'ElseIf Weekday(Data) + 1 = 2 Then
'    Data = Data + 2
'Else
'    Data = Data
'End If

i = 0

For i = 0 To 7 Step 1
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(data, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    
    If rsFer.RecordCount = 0 Then
        rsFer.Close
        Exit For
    End If
    
    If rsFer.RecordCount <> 0 Then
    data = data + 1
      If Weekday(data) = 1 Then
        data = data + 1
      End If
      If Weekday(data) = 7 Then
        data = data + 2
      End If
    End If
    rsFer.Close
Next
'Set db = Nothing


'SeImed(DiaSem([DTVcto_LctoFin]+1)=7;[DTVcto_LctoFin]+3;
'SeImed(DiaSem([DTVcto_LctoFin]+1)=1;[DTVcto_LctoFin]+3;
'SeImed(DiaSem([DTVcto_LctoFin]+1)=2;[DTVcto_LctoFin]+2;
'[DTVcto_LctoFin]+1)

CalculaDataCob = data

End Function


'Public Function GetUserLevel() As Long
'Dim gp As DAO.Group
'Dim lngCurLevel As Long
'Dim lngLevel As Long
'
'For Each gp In Workspaces(0).Users(CurrentUser()).Groups
'    Select Case gp.Name
'        Case "Full Permissions"
'            lngCurLevel = 3
'        Case "Fulll Data Users"
'            lngCurLevel = 2
'        Case "Users"
'            lngCurLevel = 1
'    End Select
'    lngLevel = IIf(lngLevel > lngCurLevel, lngLevel, lngCurLevel)
'Next gp
'
'GetUserLevel = lngLevel
'
'Set gp = Nothing
'
'End Function

Public Function GetUserGroup()

Dim wrkDefault As Workspace
    Dim usrNew As User
    Dim usrLoop As User
    Dim grpNew As Group
    Dim grpLoop As Group
    Dim grpMember As Group

    Set wrkDefault = DBEngine.Workspaces(0)

    With wrkDefault

        ' Cria e acrescenta o novo usu�rio.
        'Set usrNew = .CreateUser("Francisco Silva", _
        '    "abc123DEF456", "Senha1")
        '.Users.Append usrNew

        ' Cria e acrescenta o novo grupo.
        'Set grpNew = .CreateGroup("Contas", _
        '    "UVW987xyz654")

'.Groups.Append grpNew

        ' Torne o usu�rio Francisco Silva um membro do
        ' grupo Contas criando e adicionando o objeto
        ' Group adequado � cole��o Groups do usu�rio.
        ' O mesmo � conseguido se um objeto User
        ' que represente Francisco Silva for criado
        ' e acrescentado � cole��o Users do grupo
        ' Contas.
        'Set grpMember = usrNew.CreateGroup("Contas")
        'usrNew.Groups.Append grpMember

        'Debug.Print "Cole��o Users:"

        ' Enumera todos os objetos User na cole��o

' Users do espa�o de trabalho padr�o.
        For Each usrLoop In .Users
            Debug.Print "    " & usrLoop.Name
            Debug.Print "        Pertence a estes grupos:"

            ' Enumera todos os objetos Group em cada
            ' cole��o Groups do objeto User.
            If usrLoop.Groups.count <> 0 Then
                For Each grpLoop In usrLoop.Groups
                    Debug.Print "            " & _
                        grpLoop.Name
                Next grpLoop
            Else
                Debug.Print "            [Nenhum]"

End If

        Next usrLoop

        Debug.Print "Cole��o Groups:"

        ' Enumera todos os objetos Group na cole��o
        ' Groups do espa�o de trabalho padr�o.
        For Each grpLoop In .Groups
            Debug.Print "    " & grpLoop.Name
            Debug.Print "        Tem como seus membros:"

            ' Enumera todos os objetos User em cada
            ' cole��o Users do objeto Group.
            If grpLoop.Users.count <> 0 Then
                For Each usrLoop In grpLoop.Users
                    Debug.Print "            " & usrLoop.Name
                Next usrLoop
            Else
                Debug.Print "            [Nenhum]"
            End If

        Next grpLoop

        ' Exclui os objetos User e Group pois isto �
        ' somente uma demonstra��o.
        '.Users.Delete "Francisco Silva"
        '.Groups.Delete "Contas"

    End With

End Function


Public Function AbreSelArquivo(ByVal frmHwnd As Single, Optional Titulo As String = "Sele��o de arquivos", Optional DiretorioInicial As String = "C:\", Optional Filtro As String) As String
On Error GoTo Err_AbreSelArquivo
Dim hWnd
'API do Windows que mostra a janela de abertura de arquivo

' BACKUP DO PADR�O DE FILTRO ====>>>> "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    
    'Esta execu��o pode ser ignorada se der erro
    OFName.hwndOwner = hWnd
    
    
    If IsMissing(Filtro) Or Trim$(Filtro) = "" Then
     'Filtro de Arquivos
     OFName.lpstrFilter = "Imagens v�lidas" + Chr$(0) + "*.png;*.bmp;*.jpg;*.xls" + Chr$(0)

    Else
     OFName.lpstrFilter = Filtro
    End If

    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = DiretorioInicial
    OFName.lpstrTitle = Titulo
    
    OFName.Flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        AbreSelArquivo = Trim$(OFName.lpstrFile)
    Else
        AbreSelArquivo = ""
    End If
    
Exit_AbreSelArquivo:
    Exit Function
Err_AbreSelArquivo:
    MsgBox Err.Description
    Resume Exit_AbreSelArquivo
End Function


Public Function OpenGetFileDialog( _
    Optional ByRef DialogTitle As Variant, _
    Optional ByVal InitialDir As Variant, _
    Optional ByVal Filter As Variant _
) As String

On Error GoTo Err_OpenGetFileDialog

    Dim OpenFile    As OPENFILENAME
    Dim lReturn     As Long

    If IsMissing(DialogTitle) Then DialogTitle = "Default Title"
    If IsMissing(InitialDir) Then InitialDir = "C:\"
    If IsMissing(Filter) Then Filter = ""
    
    OpenFile.lStructSize = LenB(OpenFile)
    OpenFile.hwndOwner = Application.hWndAccessApp
    OpenFile.lpstrFile = String(256, 0)
    OpenFile.nMaxFile = LenB(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = InitialDir
    OpenFile.lpstrFilter = Filter
    OpenFile.lpstrTitle = DialogTitle
    OpenFile.Flags = 0
    
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        OpenGetFileDialog = ""
    Else
        OpenGetFileDialog = OpenFile.lpstrFile
    End If
Exit_OpenGetFileDialog:
    Exit Function
Err_OpenGetFileDialog:
    MsgBox Err.Description
    Resume Exit_OpenGetFileDialog
End Function
Function STRMaiuscula(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim x
  Dim nova As String
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "u"
      Case "�", "�"
        x = "c"
      Case "�", "�"
        x = "."
      Case "!", "@", "#", "%", "%", "^", "'", "&", "*", "_", "+", "=", ":", ";", "?", ">", "<", "~", "`", "|", "\", Chr$(34)
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRMaiuscula = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function
Function STRMaiusculaSINT(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim x
  Dim nova As String
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "u"
      Case "�", "�"
        x = "c"
      Case "�", "�"
        x = ""
      Case "!", "@", "#", "%", "%", "^", "'", "&", "*", "_", "+", "=", ":", ";", "?", ">", "<", "~", "`", "|", "\", Chr$(34)
        x = " "
          Case ".", "-", "/", ",", "(", ")", "`", "~", "'", "�", "^"
        x = ""

      Case Else
      x = x
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRMaiusculaSINT = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function
Function STREspeciais(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim x
  Dim nova As String
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "�"
        x = "a"
      Case "�"
        x = "A"
      Case "�"
        x = "e"
      Case "�"
        x = "E"
      Case "�"
        x = "i"
      Case "�"
        x = "I"
      Case "�"
        x = "o"
      Case "�"
        x = "O"
      Case "�"
        x = "u"
      Case "�"
        x = "U"
      Case "�", "�"
        x = ""
      Case "^", "'", "~", "`", "|", """", "�", "!", "@", "#", "$", "%", "&"
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STREspeciais = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Function STRAcentos(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim x
  Dim nova As String
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�"
        x = "A"
      Case "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�"
        x = "E"
      Case "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�"
        x = "I"
      Case "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�"
        x = "O"
      Case "�", "�", "�", "�"
        x = "u"
      Case "�", "�", "�", "�"
        x = "U"
      Case "�"
        x = "c"
      Case "�"
        x = "C"
      Case "�", "�"
        x = "."
      Case "^", "'", "~", "`"
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRAcentos = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function
Function STRUnicode(campo As Variant) As String
  On Error GoTo Err_STR
 
'� \u00e1
'� \u00e0
'� \u00e2
'� \u00e3
'� \u00e4
'� \u00c1
'� \u00c0
'� \u00c2
'� \u00c3
'� \u00c4
campo = Replace(campo, "\u00e1", "�")
campo = Replace(campo, "\u00e0", "�")
campo = Replace(campo, "\u00e2", "�")
campo = Replace(campo, "\u00e3", "�")
campo = Replace(campo, "\u00e4", "�")
campo = Replace(campo, "\u00c1", "�")
campo = Replace(campo, "\u00c0", "�")
campo = Replace(campo, "\u00c2", "�")
campo = Replace(campo, "\u00c3", "�")
campo = Replace(campo, "\u00c4", "�")
'� \u00e9
'� \u00e8
'� \u00ea
'� \u00c9
'� \u00c8
'� \u00ca
'� \u00cb
campo = Replace(campo, "\u00e9", "�")
campo = Replace(campo, "\u00e8", "�")
campo = Replace(campo, "\u00ea", "�")
campo = Replace(campo, "\u00c9", "�")
campo = Replace(campo, "\u00c8", "�")
campo = Replace(campo, "\u00ca", "�")
campo = Replace(campo, "\u00cb", "�")
'� \u00ed
'� \u00ec
'� \u00ee
'� \u00ef
'� \u00cd
'� \u00cc
'� \u00ce
'� \u00cf
campo = Replace(campo, "\u00ed", "�")
campo = Replace(campo, "\u00ec", "�")
campo = Replace(campo, "\u00ee", "�")
campo = Replace(campo, "\u00ef", "�")
campo = Replace(campo, "\u00cd", "�")
campo = Replace(campo, "\u00cc", "�")
campo = Replace(campo, "\u00ce", "�")
campo = Replace(campo, "\u00cf", "�")
'� \u00f3
'� \u00f2
'� \u00f4
'� \u00f5
'� \u00f6
'� \u00d3
'� \u00d2
'� \u00d4
'� \u00d5
'� \u00d6
campo = Replace(campo, "\u00f3", "�")
campo = Replace(campo, "\u00f2", "�")
campo = Replace(campo, "\u00f4", "�")
campo = Replace(campo, "\u00f5", "�")
campo = Replace(campo, "\u00f6", "�")
campo = Replace(campo, "\u00d3", "�")
campo = Replace(campo, "\u00d2", "�")
campo = Replace(campo, "\u00d4", "�")
campo = Replace(campo, "\u00d5", "�")
campo = Replace(campo, "\u00d6", "�")
'� \u00fa
'� \u00f9
'� \u00fb
'� \u00fc
'� \u00da
'� \u00d9
'� \u00db
campo = Replace(campo, "\u00fa", "�")
campo = Replace(campo, "\u00f9", "�")
campo = Replace(campo, "\u00fb", "�")
campo = Replace(campo, "\u00fc", "�")
campo = Replace(campo, "\u00da", "�")
campo = Replace(campo, "\u00d9", "�")
campo = Replace(campo, "\u00db", "�")
'� \u00e7
'� \u00c7
'� \u00f1
'� \u00d1
campo = Replace(campo, "\u00e7", "�")
campo = Replace(campo, "\u00c7", "�")
campo = Replace(campo, "\u00f1", "�")
campo = Replace(campo, "\u00d1", "�")

STRUnicode = campo
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function


Public Function MMCase(texto As String) As String
Dim sPalavra As String, iPosIni As Integer
Dim iPosFim As Integer, sResultado As String
iPosIni = 1
texto = LCase(texto) & " "

Do Until InStr(iPosIni, texto, " ") = 0
iPosFim = InStr(iPosIni, texto, " ")
sPalavra = Mid(texto, iPosIni, iPosFim - iPosIni)
iPosIni = iPosFim + 1
If sPalavra <> "de" And sPalavra <> "da" And _
sPalavra <> "do" And sPalavra <> "das" _
And sPalavra <> "dos" And sPalavra <> _
"a" And sPalavra <> "e" Then
sPalavra = UCase(left(sPalavra, 1)) & _
LCase(Mid(sPalavra, 2))
End If
sResultado = sResultado & " " & sPalavra
Loop
MMCase = Trim(sResultado)
End Function


Function TextoPor2(texto As String) As Variant

Dim arr(2) As Variant
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim parte1 As String
Dim parte2 As String
Dim parte3 As String

For i = 25 To 1 Step -1

If Mid(texto, i, 1) = " " Then
    Exit For
End If
    TextoPor2 = arr
Next
    parte1 = Mid(texto, 1, i)
    
For j = (50 - (25 - i)) To 1 Step -1
If Mid(texto, j, 1) = " " Then
    Exit For
End If
    TextoPor2 = arr
Next
    parte2 = Mid(texto, i + 1, j - i)
    
For k = 75 - (50 - j) To 1 Step -1
If Mid(texto, k, 1) = " " Then
    Exit For
End If
    TextoPor2 = arr
Next
    parte3 = Mid(texto, j + 1, 75)




arr(0) = parte1
arr(1) = parte2
arr(2) = parte3

TextoPor2 = arr

End Function

Public Sub ReIniciaMySqlUpdate()
'Limpa o Vetor de Campos para atualiza��o da base
Erase Campos
End Sub

Public Sub AlterarCampo(ByVal campo As ADODB.Field, Optional ByVal valor)
On Error GoTo NoBug
Dim PrecisaAspas As Boolean
Dim CId As Integer
Dim BoundErr As Boolean

'Verificando o tipo de campo do qual ser� trabalhado

'Se o valor for nulo, deixar como nulo
'Cada tipo de campo retorna um n�mero
If IsNull(valor) Then
 PrecisaAspas = False
 valor = Null
Else
 PrecisaAspas = False
 Select Case campo.Type
 Case 2    ' Tinyint
  PrecisaAspas = False
  valor = CInt(valor)
 Case 3    ' Integer
  PrecisaAspas = False
  valor = CInt(valor)
 Case 5    ' Double
  PrecisaAspas = False
  valor = CDbl(valor)
 Case 16   ' TinyInt
  PrecisaAspas = False
  valor = CInt(valor)
 Case 19
  PrecisaAspas = False
  valor = CInt(valor)
 Case 129  ' Char
  PrecisaAspas = True
  valor = CStr(valor)
 Case 133  ' DateTime
  PrecisaAspas = True
  valor = CStr(Format(valor, "yyyy/MM/dd"))
 Case 135  ' DateTime
  PrecisaAspas = True
  If Len(valor) <= 8 Then
    valor = CStr(Format(valor, "hh:mm:ss"))
  Else
    valor = CStr(Format(valor, "yyyy/MM/dd"))
  End If
 Case 200  ' VarChar
  valor = CStr(valor)
  PrecisaAspas = True
 Case 201  ' VarChar
  PrecisaAspas = True
  valor = CStr(valor)
 Case 202  ' VarChar
  PrecisaAspas = True
  valor = CStr(valor)
 Case 205 ' BLOB
  PrecisaAspas = False
  valor = ConverteParaBlob(CStr(valor))
 End Select
End If


'Se o tipo de campo solicitar as aspas (campo texto, data, blob, etc)
If PrecisaAspas = False And Not IsNull(valor) Then
 If Trim$(valor) = "" Then valor = 0
 valor = Replace(valor, ",", ".")
End If

BoundErr = True
CId = UBound(Campos) + 1
BoundErr = False

If IsNull(valor) Then
 valor = "NULL"
Else
 If PrecisaAspas Then valor = "'" & ConverteValor(CStr(valor)) & "'"
End If

'Criando mais um campo para altera��o
ReDim Preserve Campos(CId)
Campos(CId).campo = campo.Name
Campos(CId).valor = valor

Exit Sub

NoBug:
If BoundErr Then CId = 0: Resume Next
'MsgBox Err.Description
Resume Next
End Sub

Private Function ConverteParaBlob(ByVal valor As String) As String
'('), aspas duplas ("), barra invertida (\) e NUL (o byte NULL).
Dim imgValor As String
Dim n As Integer
'Muda um arquivo bin�rio (.gif, .jpg, .exe) para texto
n = FreeFile
Open valor For Binary Access Read As #n
 imgValor = Input(LOF(n), 1)
Close #n

'Converte caracteres que anulam a string
imgValor = Replace(imgValor, "'", "\'")
'imgValor = Replace(imgValor, "\", "\\")
imgValor = Replace(imgValor, vbCr, "\r")
imgValor = Replace(imgValor, vbLf, "\n")

'MsgBox imgValor
ConverteParaBlob = imgValor
End Function

Public Function ConverteValor(txt As String) As String
'Converte caracteres que anulam a string
'txt = Replace(txt, "\", "\\")
txt = Replace(txt, "'", "''")
'txt = Replace(txt, vbCr, "\r")
'txt = Replace(txt, vbLf, "\n")
ConverteValor = txt
End Function
Public Function GerarSQLUpdate(Optional ForAddNew As Boolean = False) As String
On Error GoTo NoBug
Dim i As Integer
Dim str As String
Dim a, b As String

'Gerando a string de INSERT OU UPDATE
'Se for INSERT, gerar no formato --> (Campo1, Campo2, Campo3...) VALUES (Valor1, Valor2, Valor3)
'Se for UPDATE, gerar no formato --> Campo1=Valor1, Campo2=Valor2, Campo3=Valor3...

If Not ForAddNew Then
 For i = 0 To UBound(Campos)
  str = str & Campos(i).campo & "=" & Campos(i).valor & ", "
 Next i
 str = Mid(str, 1, Len(str) - 2)
 GerarSQLUpdate = "SET " & str
Else
 For i = 0 To UBound(Campos)
  a = a & Campos(i).campo & ", "
  b = b & Campos(i).valor & ", "
 Next i
 a = Mid(a, 1, Len(a) - 2)
 b = Mid(b, 1, Len(b) - 2)
 GerarSQLUpdate = "(" & a & ") VALUES(" & b & ")"
End If

Exit Function

NoBug:
If Err.Number = 9 Then Exit Function

End Function

Public Sub SaveSQLString(strSQL As String)
On Error Resume Next
Dim i As String

i = FreeFile
Open "C:\sam\sql_string.txt" For Output As #i
 Print #i, strSQL
Close #i

Shell "notepad C:\sam\sql_string.txt"

End Sub

Function ExistePasta(path As String)

Dim drive As String
Dim pastas As String
Dim x

path = path & "\"

If InStr(path, "\") = 1 Then ' se for pasta pela rede

        'MsgBox "pasta na rede"
        drive = drive & left(path, InStr(3, path, "\") - 1)
        'MsgBox "Servidor �: " & drive & "\"
        x = CInt(Len(path)) - CInt(Len(drive))
        pastas = right(path, x)
        
        Call GerarPasta(drive, pastas)
Else

        'MsgBox " pasta local"
        drive = drive & left(path, InStr(2, path, "\") - 1)
        'MsgBox "Unidade �: " & drive & "\"
        x = CInt(Len(path)) - CInt(Len(drive))
        pastas = right(path, x)
                     
        Call GerarPasta(drive, pastas)

End If


End Function
Sub GerarPasta(sDrive As String, sDir As String)

Dim sBuild As String

While InStr(2, sDir, "\") > 1

    sBuild = sBuild & left(sDir, InStr(2, sDir, "\") - 1) & "\"
    sDir = Mid$(sDir, InStr(2, sDir, "\"))
    
    If Dir$(sDrive & sBuild, 16) = "" Then
        MkDir sDrive & sBuild
    End If
Wend
End Sub

Function STRNome(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim x
  Dim nova As String
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "@", "#", "&", "*", "_", ":", ";", "'", "|", "\", "/"
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRNome = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Function fSearchFile(ByVal strFileName As String, _
            ByVal strSearchPath As String) As String
'Returns the first match found
    Dim lpBuffer As String
    Dim lngResult As Long
    fSearchFile = ""
    lpBuffer = String$(1024, 0)
    lngResult = apiSearchTreeForFile(strSearchPath, strFileName, lpBuffer)
    If lngResult <> 0 Then
        If InStr(lpBuffer, vbNullChar) > 0 Then
            fSearchFile = left$(lpBuffer, InStr(lpBuffer, vbNullChar) - 1)
        End If
    End If
End Function


Public Function TestarVinculosSQL() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQL = True

'Obtendo uma conex�o qualquer de uma tabela qualquer vinculada no sistema para testar o v�nculo
CurrentDb.QueryDefs.Delete "GetActualConnection"
ConnX = CurrentDb.CreateQueryDef("GetActualConnection", "SELECT Connect FROM MSysObjects WHERE Flags=537919488 AND Name like 'Cadastro de Produtos' AND NOT ISNULL(Connect);").OpenRecordset!Connect
CurrentDb.QueryDefs.Delete "GetActualConnection"

'Se a conex�o for satisfat�ria, sair da rotina
If Trim$(" " & ConnX) = "" Then TestarVinculosSQL = True: Exit Function
If TestarConexao(ConnX) Then Exit Function

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'SQL' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'SQL';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQL = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQL = False
Exit Function

NoBug:


If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function
Public Function TestarVinculosCEP() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosCEP = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'CEP' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'CEP';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosCEP = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosCEP = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function
Public Function TestarVinculosWH() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosWH = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'WH' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'WH';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosWH = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosWH = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function TestarVinculosTABSQL() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosTABSQL = True

'Obtendo uma conex�o qualquer de uma tabela qualquer vinculada no sistema para testar o v�nculo
CurrentDb.QueryDefs.Delete "GetActualConnection"
ConnX = CurrentDb.CreateQueryDef("GetActualConnection", "SELECT Connect FROM MSysObjects WHERE Flags=537919488 AND Name like 'Detalhe Produtos Vendidos_Antigo' AND NOT ISNULL(Connect);").OpenRecordset!Connect
CurrentDb.QueryDefs.Delete "GetActualConnection"

'Se a conex�o for satisfat�ria, sair da rotina
If Trim$(" " & ConnX) = "" Then TestarVinculosTABSQL = True: Exit Function
If TestarConexao(ConnX) Then Exit Function

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'TAB' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'TAB';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosTABSQL = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosTABSQL = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function
Public Function TestarVinculosSISSQL() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSISSQL = True

'Obtendo uma conex�o qualquer de uma tabela qualquer vinculada no sistema para testar o v�nculo
CurrentDb.QueryDefs.Delete "GetActualConnection"
ConnX = CurrentDb.CreateQueryDef("GetActualConnection", "SELECT Connect FROM MSysObjects WHERE Flags=537919488 AND Name like 'Cadastro de produtos_Sispedal' AND NOT ISNULL(Connect);").OpenRecordset!Connect
CurrentDb.QueryDefs.Delete "GetActualConnection"

'Se a conex�o for satisfat�ria, sair da rotina
If Trim$(" " & ConnX) = "" Then TestarVinculosSISSQL = True: Exit Function
If TestarConexao(ConnX) Then Exit Function

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'SIS' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'SIS';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSISSQL = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSISSQL = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function



Public Function TestarConexao(StringConexao As String) As Boolean
On Error GoTo NoBug
Dim ConnTest As New ADODB.Connection

'SEMPRE DEIXAR ESTE VALOR COMO TRUE
TestarConexao = True

'Abrindo a conex�o para verificar se est� correta.
ConnTest.Open StringConexao
ConnTest.Close

Exit Function
NoBug:
'Caso ocorra algum erro, a conex�o n�o � v�lida
TestarConexao = False
Exit Function
End Function

Function Trunca(dblNumero As Double, dblDecimais As Double) As Double
    dblNumero = dblNumero * 10 ^ dblDecimais
    If Mid(right(dblNumero, 2), 1, 1) = "," Or Mid(right(dblNumero, 3), 1, 1) = "," Or Mid(right(dblNumero, 4), 1, 1) = "," Then
        Trunca = Fix(dblNumero) / 10 ^ dblDecimais
    Else
        Trunca = (dblNumero) / 10 ^ dblDecimais
    End If
End Function
Function TruncaIPI(dblNumero As Double, dblDecimais As Double) As Double
    dblNumero = dblNumero * 10 ^ dblDecimais
    If Mid(right(dblNumero, 2), 1, 1) = "," Or Mid(right(dblNumero, 3), 1, 1) = "," Or Mid(right(dblNumero, 4), 1, 1) = "," Or Mid(right(dblNumero, 5), 1, 1) = "," Or Mid(right(dblNumero, 6), 1, 1) = "," Or Mid(right(dblNumero, 7), 1, 1) = "," Or Mid(right(dblNumero, 8), 1, 1) = "," Or Mid(right(dblNumero, 9), 1, 1) = "," Or Mid(right(dblNumero, 10), 1, 1) = "," Or Mid(right(dblNumero, 11), 1, 1) = "," Or Mid(right(dblNumero, 12), 1, 1) = "," Then
        TruncaIPI = Fix(dblNumero) / 10 ^ dblDecimais
    Else
        TruncaIPI = (dblNumero) / 10 ^ dblDecimais
    End If
End Function
Public Function Email_CDO(Remetente As String, Destinatario As String, _
assunto As String, Corpo As String, Optional CC As String, _
Optional BCC As String, Optional Anexo1 As String, Optional Anexo2 As String)
On Error GoTo Err_STR


Dim iMsg As Object
Dim iConf As Object
Dim strBody As String
Dim Flds As Variant
Dim rsConfig As New ADODB.Recordset
Dim str As String
            
Dim emailRemetente As String
Dim nomeRemetente As String
Dim emailBcc As String
Dim arquivos As String
Dim smtpCliente As String
Dim smtpPorta As String
Dim smtpSSL As String
Dim smtpUsuario As String
Dim smtpSenha As String

AbrirConexao

If txtEnvioBol = 1 Then
    emailRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolUsuario'")
    nomeRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolNomeEmail'")
    'emailDestinatario = rsCob!Email
    emailBcc = BCC
    assunto = assunto
    smtpCliente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSMTP'")
    smtpPorta = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolPorta'")
    smtpSSL = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSSL'")
    smtpUsuario = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolUsuario'")
    smtpSenha = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSenha'")
    
ElseIf txtEnvioBol = 2 Then
    emailRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecEmail'")
    nomeRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecNomeEmail'")
    'emailDestinatario = rsCob!Email
    emailBcc = ""
    assunto = assunto
    smtpCliente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecSMTP'")
    smtpPorta = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecPorta'")
    smtpSSL = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecSSL'")
    smtpUsuario = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecUsuario'")
    smtpSenha = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecSenha'")
ElseIf txtEnvioBol = 3 Then
    emailRemetente = DLookup("[EmailUser]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")
    nomeRemetente = Remetente
    emailBcc = ""
    assunto = assunto
    smtpCliente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSMTP'")
    smtpPorta = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolPorta'")
    smtpSSL = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSSL'")
    smtpUsuario = DLookup("[EmailUser]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")
    smtpSenha = DLookup("[EmailSenha]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")

End If

 
If emailRemetente = "" Then
    MsgBox "E-mail n�o enviado, verifique configura��es do Remetente e Destinat�rio", vbInformation, "Aten��o"
    Exit Function
End If

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
 
        iConf.Load -1    ' CDO Source Defaults
        Set Flds = iConf.Fields
        With Flds
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPorta
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpCliente
            .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
            .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = smtpUsuario
            .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtpSenha
            .item("http://schemas.microsoft.com/cdo/configuration/sendemailaddress") = emailRemetente
            .Update
        End With
 
    strBody = Corpo
    With iMsg
        Set .Configuration = iConf
        .To = Destinatario
        .CC = CC
        If txtEnvioBol = 1 Then
            .ReplyTo = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolUsuario'")
            .BCC = emailBcc
        ElseIf txtEnvioBol = 2 Then
            .ReplyTo = "financeiro@proparts.esp.br" & IIf(BCC = "", "", ";" & BCC)
        ElseIf txtEnvioBol = 3 Then
            .ReplyTo = DLookup("[EmailUser]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")
            .BCC = emailBcc
        End If
        .FROM = "" & nomeRemetente & " <" & emailRemetente & ">"
        .Subject = assunto
        .TextBody = strBody
        If Anexo1 <> "" Then
            .Addattachment Anexo1
        End If
        If Anexo2 <> "" Then
            .Addattachment Anexo2
        End If
        .send
    End With
 
    Set iMsg = Nothing
    Set iConf = Nothing

Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR

End Function

Function STRArroba(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim nova As Double
  Dim x
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "@"
        x = 1
      Case Else
      x = 0
    End Select
    nova = nova + x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRArroba = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Function Est�Aberto(NomeFormul�rio As String) As Integer
On Error GoTo Err_Est�Aberto
Dim x As Integer
Dim N�meroDeFormul�rios
    N�meroDeFormul�rios = Forms.count                                            ' formul�rios.
    For x = 0 To N�meroDeFormul�rios - 1
        If Forms(x).Name = NomeFormul�rio Then
            Est�Aberto = -1
            Exit Function
        Else
            Est�Aberto = 0
        End If
    Next x
Exit_Est�Aberto:
    Exit Function
Err_Est�Aberto:
    MsgBox Error$
    Resume Exit_Est�Aberto
End Function

Public Function FSelDir(Optional InitialDir As String = "C:\") As String
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo

'API do Windows que exibe a janela de sele��o de diret�rios

With udtBI
       '.lpszTitle = lstrcat("C:\a&m", "Teste")
       .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = left$(sPath, iNull - 1)
        End If
    End If

    FSelDir = Trim$(" " & Replace(sPath, Chr(0), ""))
End Function

Function SendKeysNovo(txtTecla) As String
Dim ws As Object
Set ws = CreateObject("Wscript.shell")
ws.SendKeys txtTecla
Set ws = Nothing
End Function

Function STRTiraAcentos(campo As Variant) As String
  On Error GoTo Err_STR
  Dim a As Integer
  Dim x
  Dim nova As String
  a = 1
  x = Mid(campo, a, 1)
  While (a <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�"
        x = "A"
      Case "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�"
        x = "E"
      Case "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�"
        x = "I"
      Case "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�"
        x = "O"
      Case "�", "�", "�", "�"
        x = "u"
      Case "�", "�", "�", "�"
        x = "U"
      Case "�"
        x = "c"
      Case "�"
        x = "C"
      Case "!", "@", "#", "%", "^", "&", "_", "~", "`", "\", "�", Chr$(34), "�", "�"
      x = " "
      Case Else
      x = (x)
    End Select
    nova = nova & x
    a = a + 1
    If (a <= Len(campo)) Then
      x = Mid(campo, a, 1)
    End If
  Wend
  STRTiraAcentos = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Public Sub RegistrarErro(Optional Numero As Double, Optional Descricao As String, Optional Modulo As String, Optional Funcao As String)
Dim MyPath As String
Dim parts() As String
Dim i As Integer
Dim ErrorFile As String
Dim FN As Integer
Dim strError As String

'Gera o arquivo de log de erro em A&M_SIS_Diretorio\pml_err.log

strError = "Log de erro gerado em : " & Format(Now, "dd/mm/yyyy") & " �s " & Format(Now, "hh:mm:ss") & vbCrLf
strError = strError & "M�dulo : " & Modulo & vbCrLf & "Fun��o : " & Funcao & vbCrLf
strError = strError & "N�mero : " & Numero & vbCrLf & "Descri��o : " & Descricao & vbCrLf & vbCrLf

MyPath = CurrentDb.Name

parts = Split(MyPath, "\")

ErrorFile = ""
For i = 0 To UBound(parts) - 1
 ErrorFile = ErrorFile & parts(i) & "\"
Next i
ErrorFile = ErrorFile & "Sisparts_err.log"

FN = FreeFile
Open ErrorFile For Append As #FN
 Print #FN, strError
Close #FN


End Sub

Function FindAndReplace(ByVal strInString As String, _
        strFindString As String, _
        strReplaceString As String) As String
Dim intPtr As Integer
    If Len(strFindString) > 0 Then  'catch if try to find empty string
        Do
            intPtr = InStr(strInString, strFindString)
            If intPtr > 0 Then
                FindAndReplace = FindAndReplace & left(strInString, intPtr - 1) & _
                                        strReplaceString
                    strInString = Mid(strInString, intPtr + Len(strFindString))
            End If
        Loop While intPtr > 0
    End If
    FindAndReplace = FindAndReplace & strInString
End Function

Function Fix6(x As Variant) As Currency
On Error GoTo Err_Fix6
    Fix6 = Int(Nz(x) * 1000000 + 0.5) / 1000000
Exit_Fix6:
    Exit Function
Err_Fix6:
    MsgBox Error$
    Resume Exit_Fix6
End Function

Public Function TestarVinculosSQLSP() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQLSP = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'NFESP' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'NFESP';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQLSP = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQLSP = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function TestarVinculosSQLES() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQLES = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'NFEES' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'NFEES';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQLES = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQLES = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function TestarVinculosSQLSC() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQLSC = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'NFESC' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'NFESC';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQLSC = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQLSC = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function UserAltStatusCob(txtID As Long, txtAlt As String, txtIDStatusCob As Long, txtAberto As String)
Dim txtDescrTPStatusCob As String
Dim txtMsg As String
AbrirConexao
If IsNull(txtAlt) Or txtAlt = "" Then
    txtDescrTPStatusCob = DLookup("[Descr_TPStatus]", "tblTPStatus", "[ID_TPStatus] = " & txtIDStatusCob & "")
    txtMsg = Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm") & " " & txtDescrTPStatusCob
    CNN.Execute "UPDATE tblLctoFin SET " _
    & "tblLctoFin.UserAlt_LctoFin = '" & txtMsg & "' " _
    & "WHERE (((tblLctoFin.ID_LctoFin)= " & txtID & "));", dbSeeChanges
Else
    txtDescrTPStatusCob = DLookup("[Descr_TPStatus]", "tblTPStatus", "[ID_TPStatus] = " & txtIDStatusCob & "")
    txtMsg = txtAlt & "->" & Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm") & " " & txtDescrTPStatusCob
    CNN.Execute "UPDATE tblLctoFin SET " _
    & "tblLctoFin.UserAlt_LctoFin = '" & txtMsg & "' " _
    & "WHERE (((tblLctoFin.ID_LctoFin)= " & txtID & "));", dbSeeChanges
End If
If txtAberto = "S" Then
    Forms!frmLctoFin_CntRec_Cadastro!UserAlt_LctoFin = txtMsg
End If
End Function

Public Function fncDiasUteis(dataLan�amento As Date, DataRef As Date) As Integer
Dim j%, dataAnalisada As Date
Dim sqlString As String
Dim rsFer As New ADODB.Recordset
Dim txtFeriado As Boolean

dataAnalisada = dataLan�amento '+ 1
AbrirConexao
Do While Not dataAnalisada > DataRef
    txtFeriado = False
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(dataAnalisada, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    If rsFer.RecordCount = 0 Then
        rsFer.Close
    Else
        txtFeriado = True
        rsFer.Close
    End If

    If Eval("weekday(#" & Format(dataAnalisada, "mm/dd/yyyy") & "#) between 2 and 6") And txtFeriado = False Then
            j = j + 1
    End If
    dataAnalisada = dataAnalisada + 1
Loop
fncDiasUteis = j
End Function


<<<<<<< HEAD:referencia/code/modFuncoesGerais.bas
=======
=======
Attribute VB_Name = "modFuncoesGerais"
Option Compare Database
Global Var_Acesso As Integer
Global txtEnvioBol As Integer
Global frmNomeForm As String
Global IDVDAtu As Long
'############################################
'######### INICIO DAS APIs DO WINDOWS #######
'############################################
'######## ==>>> API DE SELE��O DE ARQUIVOS! #####################################
'######## ==>>>>>>>> Dependente da DLL COMDLG32.DLL - REQUER W95 OU SUPERIOR !


'Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Type OPENFILENAME
'    lStructSize As Long
'    hwndOwner As Long
'    hInstance As Long
'    lpstrFilter As String
'    lpstrCustomFilter As String
'    nMaxCustFilter As Long
'    nFilterIndex As Long
'    lpstrFile As String
'    nMaxFile As Long
'    lpstrFileTitle As String
'    nMaxFileTitle As Long
'    lpstrInitialDir As String
'    lpstrTitle As String
'    Flags As Long
'    nFileOffset As Integer
'    nFileExtension As Integer
'    lpstrDefExt As String
'    lCustData As Long
'    lpfnHook As Long
'    lpTemplateName As String
'End Type
#If VBA7 Then
    Private Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare PtrSafe Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
    Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

#If VBA7 Then
    Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As LongPtr
    lpTemplateName As String
    End Type
#Else
    Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    End Type
#End If

'###############################################################################

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260

'Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi As BrowseInfo) As Long
'Private Declare Function SHGetPathFromIDList Lib "Shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'
'Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
'        "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
'        As String, ByVal lpOutPath As String) As Long

#If VBA7 Then
    Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Public Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Public Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Public Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    
    Public Declare PtrSafe Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
            "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
            As String, ByVal lpOutPath As String) As Long

#Else
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    
    Private Declare Function apiSearchTreeForFile Lib "ImageHlp.dll" Alias _
            "SearchTreeForFile" (ByVal lpRoot As String, ByVal lpInPath _
            As String, ByVal lpOutPath As String) As Long
#End If


'############################################
'######### FINAL DAS  APIs DO WINDOWS #######
'############################################
Public Type T_MySql
 campo As String
 valor As Variant
End Type

Public Enum T_Valor
 Num�rico = 0
 texto = 1
 Auto_Detec��o = 2
End Enum

Dim MySqlConn As ADODB.Connection
Dim LCampos As String
Dim Campos() As T_MySql
Global ulibera As String

'Permiss�es
Global txtPer
Option Explicit

'*****************************************************
#If VBA7 Then
Const ChaveReg = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\Common\Theme"
#Else
Const ChaveReg = "HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Common\Theme"
#End If
 
Enum eCor
   Azul = 1
   Prata = 2
   Preto = 3
End Enum
'*****************************************************


Function EstaAbertoForm(NomeFormul�rio As String) As Integer
On Error GoTo Err_Est�Aberto
Dim N�meroDeFormul�rios
Dim x As Integer
    N�meroDeFormul�rios = Forms.count                                            ' formul�rios.
    For x = 0 To N�meroDeFormul�rios - 1
        If Forms(x).Name = NomeFormul�rio Then
            EstaAbertoForm = -1
            Exit Function
        Else
            EstaAbertoForm = 0
        End If
    Next x
Exit_Est�Aberto:
    Exit Function
Err_Est�Aberto:
    MsgBox Error$
    Resume Exit_Est�Aberto
End Function

Public Sub fncCorAccess(cor As eCor)
Dim ChaveReg As String
Dim objReg As Object: Set objReg = CreateObject("wscript.shell")
'Gravar valor na chave THEME
objReg.RegWrite ChaveReg, cor, "REG_DWORD"
Set objReg = Nothing
End Sub
'Para alterar a cor , basta executar o procedimento: Call fncCorAccess(Preto)
'Reiniciando o Access
'O �nico por�m � que a cor s� � alterada na abertura seguinte do Access.  Para criar um efeito imediato na configura��o, montei um c�digo que reinicia automaticamente a aplica��o. Aqui est�:

Public Sub fncReiniciandoAplicativo()
Dim strLocal As String
Dim objWs As Object
'------------------------
'Reiniciando
'------------------------
Set objWs = CreateObject("wscript.shell")
strLocal = CurrentProject.path & "\" & CurrentProject.Name
strLocal = Chr(34) & "MSACCESS.EXE" & Chr(34) & " " & Chr(34) _
           & strLocal & Chr(34)
'0 - oculto / 5 - vis�vel
objWs.Run strLocal, 5, "false"
'-------------------------------------
'Fecha atual
'-------------------------------------
Application.Quit acQuitSaveAll
End Sub


Function AbrirCombo()
On Error GoTo Err_AbrirCombo
    SendKeysNovo ("{F4}")
Exit_AbrirCombo:
    Exit Function
Err_AbrirCombo:
    MsgBox Error$
    Resume Exit_AbrirCombo
End Function

Function AbrirFormul�rio(NomeFormul�rio As String, Crit�rio As String) As Integer
On Error GoTo Err_AbrirFormul�rio
    DoCmd.OpenForm NomeFormul�rio, , , Crit�rio
Exit_AbrirFormul�rio:
    Exit Function
Err_AbrirFormul�rio:
    MsgBox Error$
    Resume Exit_AbrirFormul�rio
End Function

Function AbrirObj(varTPObj As Variant, varNmObj As Variant) As Integer
On Error GoTo Err_AbrirObj
    Select Case varTPObj
    Case "Formul�rio"
        DoCmd.OpenForm varNmObj
    Case "Relat�rio"
        DoCmd.OpenReport varNmObj, A_PREVIEW
    End Select
Exit_AbrirObj:
    Exit Function
Err_AbrirObj:
    MsgBox Error
    Resume Exit_AbrirObj
End Function

Function Acesso() As Integer
On Error GoTo Err_Acesso
Dim db  As Database, rs As Recordset, x As Long, Doc As Document
Dim y
Dim i
  
    Set db = DBEngine.Workspaces(0).Databases(0)
    Set rs = db.OpenRecordset("tblAM_Acesso")
    Set Doc = db.Containers(7).Documents("tblAM_Acesso")
    x = 0
    While Not rs.EOF
        x = x + 1
        rs.MoveNext
    Wend
    If x < 30 Then
        y = MsgBox("Esta � a " & (x + 1) & "� vez que voc� acessa a demonstra��o deste produto!", 64, "Aten��o")
        y = AbrirFormul�rio("frmIni", "")
        rs.AddNew
        rs!DT_Acesso = Date
        rs.Update
        Var_Acesso = 1
    If Doc.UserName = "SUPERVISOR" Then
      If MsgBox("Deseja liberar o acesso?", 36, "Libera��o do Sistema") = 6 Then
        For i = x To 31
          rs.AddNew
          rs!DT_Acesso = Date
          rs.Update
        Next i
      End If
    End If
    Else
        If x = 30 Then
            y = MsgBox("Esta � a " & (x + 1) & "� vez que voc� acessa a demonstra��o deste produto!", 48, "Aten��o")
            If Doc.UserName = "SUPERVISOR" Then
              MsgBox Trim(Mid(str$(CLng(Date)), 1, 3) & Chr(Day(Date) / 2 + 65) & Mid(str$(CLng(Date)), 4, 4))
            End If
            Var_Acesso = Senha()
        Else
            Var_Acesso = 1
        End If
    End If
    rs.Close
    Acesso = Var_Acesso
Exit_Acesso:
    Exit Function
Err_Acesso:
    MsgBox Error$
    Resume Exit_Acesso
End Function

Function Ajuda()
On Error GoTo Err_Ajuda
    SendKeysNovo ("{F1}")
Exit_Ajuda:
    Exit Function
Err_Ajuda:
    MsgBox Error$
    Resume Exit_Ajuda
End Function

Function AplicarFiltro(NomeFiltro As String) As Integer
'Aplica um filtro salvo como consulta
'Domingos Taffarello - 19/10/96
On Error GoTo Err_AplicarFiltro
    DoCmd.ApplyFilter NomeFiltro
Exit_AplicarFiltro:
    Exit Function
Err_AplicarFiltro:
    MsgBox Error$
    Resume Exit_AplicarFiltro
End Function

Function AvisarSa�da()
On Error GoTo Err_AvisarSa�da
Dim x As Integer
    If MsgBox("Deseja sair deste sistema?", 276, "Aviso de sa�da") = 6 Then
        x = Sair()
    End If
Exit_AvisarSa�da:
    Exit Function
Err_AvisarSa�da:
    MsgBox Error$
    Resume Exit_AvisarSa�da
End Function

Function ChecarAcesso()
'If (Var_Acesso = 0 And DCount("ID_Acesso", "tblAM_Acesso") >= 30) Or IsNull(Var_Acesso) Then X = Sair()
End Function

Function CompletaString(varTexto As String, strCompleta As String, intLargura As Integer, strLado As String) As String
'Completa a string varTexto com o caracter strCompleta, o n�mero de vezes at�
'atingir o comprimento intLargura. strLado indica se ser� � esquerda ou � direita
'Domingos Taffarello - 31/08/96
On Error GoTo Err_CompletaString
Dim strTemp As String
    strTemp = Trim((varTexto))
    While Len(strTemp) < intLargura
        If strLado = "E" Then
            strTemp = strCompleta & strTemp
        Else
            strTemp = strTemp & strCompleta
        End If
    Wend
    CompletaString = strTemp
Exit_CompletaString:
    Exit Function
Err_CompletaString:
    MsgBox Error$
    Resume Exit_CompletaString
End Function

Function EAN13(x) As String
On Error GoTo Err_EAN13
Dim p
Dim i
     If Len(Trim$(x)) < 12 Then
        MsgBox "Foi digitado um n�mero com menos de 12 d�gitos!"
        EAN13 = x
        Exit Function
     End If
     If Len(Trim$(x)) > 12 Then
        x = Mid(Trim$(x), 1, 12)
'        MsgBox "Foi digitado um n�mero com mais de 12 d�gitos!"
'        EAN13 = X
'        Exit Function
    End If
    p = 0
    For i = 1 To 12
        p = p + Val(Mid(Trim$(x), i, 1) * IIf(i = 1, 1, IIf((i Mod 2) = 0, 3, 1)))
    Next i
    EAN13 = x & Trim(str(((p \ 10 + IIf((p Mod 10) > 0, 1, 0)) * 10) - p))
Exit_EAN13:
    Exit Function
Err_EAN13:
    MsgBox Error$
    Resume Exit_EAN13
End Function

Function Excluir()
On Error GoTo Err_Excluir
    SendKeysNovo ("^{-}")
Exit_Excluir:
    Exit Function
Err_Excluir:
    MsgBox Error$
    Resume Exit_Excluir
End Function

Function ExibeSenha()
    MsgBox Trim(Mid(str$(CLng(Date)), 1, 3) & Chr(Day(Date) / 2 + 65) & Mid(str$(CLng(Date)), 4, 4))
End Function

Function Extenso(nValor)
'* Extenso()
'* Sintaxe..: Extenso(nValor) -> cExtenso
'* Descri��o: Retorna uma s�rie de caracteres contendo a forma extensa
'*            do valor passado como argumento.
'* Autoria..: Eng. Cesar Costa e Dalicio Guiguer Filho
'* Linguagem: Access Basic
'* Data.....: Fevereiro/1994
On Error GoTo Err_Extenso
'Faz a valida��o do argumento
  If IsNull(nValor) Or nValor <= 0 Or nValor > 9999999.99 Then
    Exit Function
  End If
'Declara as vari�veis da fun��o
  Dim nContador, nTamanho As Integer
  Dim cValor, cParte, cFinal As String
  ReDim aGrupo(4), aTexto(4) As String
'Define matrizes com extensos parciais
  ReDim aUnid(19) As String
  aUnid(1) = "UM ": aUnid(2) = "DOIS ": aUnid(3) = "TRES "
  aUnid(4) = "QUATRO ": aUnid(5) = "CINCO ": aUnid(6) = "SEIS "
  aUnid(7) = "SETE ": aUnid(8) = "OITO ": aUnid(9) = "NOVE "
  aUnid(10) = "DEZ ": aUnid(11) = "ONZE ": aUnid(12) = "DOZE "
  aUnid(13) = "TREZE ": aUnid(14) = "QUATORZE ": aUnid(15) = "QUINZE "
  aUnid(16) = "DEZESSEIS ": aUnid(17) = "DEZESSETE ": aUnid(18) = "DEZOITO "
  aUnid(19) = "DEZENOVE "
  ReDim aDezena(9) As String
  aDezena(1) = "DEZ ": aDezena(2) = "VINTE ": aDezena(3) = "TRINTA "
  aDezena(4) = "QUARENTA ": aDezena(5) = "CINQUENTA "
  aDezena(6) = "SESSENTA ": aDezena(7) = "SETENTA ": aDezena(8) = "OITENTA "
  aDezena(9) = "NOVENTA "
  ReDim aCentena(9) As String
  aCentena(1) = "CENTO ":  aCentena(2) = "DUZENTOS "
  aCentena(3) = "TREZENTOS ": aCentena(4) = "QUATROCENTOS "
  aCentena(5) = "QUINHENTOS ": aCentena(6) = "SEISCENTOS "
  aCentena(7) = "SETECENTOS ": aCentena(8) = "OITOCENTOS "
  aCentena(9) = "NOVECENTOS "
'Divide o valor em v�rios grupos
  cValor = Format$(nValor, "0000000000.00")
  aGrupo(1) = Mid$(cValor, 2, 3)
  aGrupo(2) = Mid$(cValor, 5, 3)
  aGrupo(3) = Mid$(cValor, 8, 3)
  aGrupo(4) = "0" + Mid$(cValor, 12, 2)
'Processa cada grupo
  For nContador = 1 To 4
    cParte = aGrupo(nContador)
    nTamanho = Switch(Val(cParte) < 10, 1, Val(cParte) < 100, 2, Val(cParte) < 1000, 3)
    If nTamanho = 3 Then
      If right$(cParte, 2) <> "00" Then
        aTexto(nContador) = aTexto(nContador) + aCentena(left(cParte, 1)) + "E "
        nTamanho = 2
      Else
        aTexto(nContador) = aTexto(nContador) + IIf(left$(cParte, 1) = "1", "CEM ", aCentena(left(cParte, 1)))
      End If
    End If
    If nTamanho = 2 Then
      If Val(right(cParte, 2)) < 20 Then
        aTexto(nContador) = aTexto(nContador) + aUnid(right(cParte, 2))
      Else
        aTexto(nContador) = aTexto(nContador) + aDezena(Mid(cParte, 2, 1))
        If right$(cParte, 1) <> "0" Then
          aTexto(nContador) = aTexto(nContador) + "E "
          nTamanho = 1
        End If
      End If
    End If
    If nTamanho = 1 Then
      aTexto(nContador) = aTexto(nContador) + aUnid(right(cParte, 1))
    End If
  Next
'Gera o formato final do texto
  If Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 0 And Val(aGrupo(4)) <> 0 Then
    cFinal = aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS")
  Else
    cFinal = ""
    cFinal = cFinal + IIf(Val(aGrupo(1)) <> 0, aTexto(1) + IIf(Val(aGrupo(1)) > 1, "MILH�ES ", "MILH�O "), "")
    If Val(aGrupo(2) + aGrupo(3)) = 0 Then
      cFinal = cFinal + "DE "
    Else
      cFinal = cFinal + IIf(Val(aGrupo(2)) <> 0, aTexto(2) + "MIL ", "")
    End If
    cFinal = cFinal + aTexto(3) + IIf(Val(aGrupo(1) + aGrupo(2) + aGrupo(3)) = 1, "REAL ", "REAIS ")
    cFinal = cFinal + IIf(Val(aGrupo(4)) <> 0, "E " + aTexto(4) + IIf(Val(aGrupo(4)) = 1, "CENTAVO", "CENTAVOS"), "")
  End If
  Extenso = cFinal
Exit_Extenso:
    Exit Function
Err_Extenso:
    MsgBox Error$
    Resume Exit_Extenso
End Function

'Function Fechar()
'On Error GoTo Err_Close
'    DoCmd.Close
'Exit_Close:
'    Exit Function
'Err_Close:
'    MsgBox Error$
'    Resume Exit_Close
'End Function

Function fncFechar(NomeFormul�rio As String)
On Error GoTo Err_FecharFormul�rio
    DoCmd.Close A_FORM, NomeFormul�rio
    'DoCmd.Close
Exit_FecharFormul�rio:
    Exit Function
Err_FecharFormul�rio:
    MsgBox Error$
    Resume Exit_FecharFormul�rio
End Function

Function Fix2(x As Variant) As Currency '' #AILTON - ARREDONDAMENTO
On Error GoTo Err_Fix2
    Fix2 = Int(Nz(x) * 100 + 0.5) / 100
Exit_Fix2:
    Exit Function
Err_Fix2:
    MsgBox Error$
    Resume Exit_Fix2
End Function
Function Fix4(x As Variant) As Currency
On Error GoTo Err_Fix4
    Fix4 = Int(Nz(x) * 10000 + 0.5) / 10000
Exit_Fix4:
    Exit Function
Err_Fix4:
    MsgBox Error$
    Resume Exit_Fix4
End Function


Function ForaDaLista()
On Error GoTo Err_ForaDaLista
Dim Response
    Response = DATA_ERRCONTINUE
    Screen.ActiveControl = Null
    SendKeysNovo ("{ESC}")
Exit_ForaDaLista:
    Exit Function
Err_ForaDaLista:
    MsgBox Error$
    Resume Exit_ForaDaLista
End Function

Function ImprimirRelat�rio(NomeRelat�rio As String, Crit�rio As String) As Integer
On Error GoTo Err_ImprimeNota_Click
    DoCmd.OpenReport NomeRelat�rio, A_PREVIEW, , Crit�rio
Exit_ImprimeNota_Click:
    Exit Function

Err_ImprimeNota_Click:
    MsgBox Error$
    Resume Exit_ImprimeNota_Click
End Function

Function Incluir(NomeForm As String, NomeControle As String)
On Error GoTo Err_Incluir
    DoCmd.OpenForm NomeForm
    Forms(NomeForm)(NomeControle).SetFocus
    DoCmd.GoToRecord , , A_NEWREC
Exit_Incluir:
    Exit Function
Err_Incluir:
    MsgBox Error$
    Resume Exit_Incluir
End Function

Function Inicializar()
'Fun��o de Inicializa��o
'Domingos Taffarello - 02/10/96
On Error GoTo Err_Inicializar

Dim x As Integer

DoCmd.OpenForm "frmLogin"
    
Exit_Inicializar:
    Exit Function
Err_Inicializar:
    MsgBox Error$
    Resume Next
    Resume Exit_Inicializar
End Function
Function IrParaControle(NomeControle As control) As Integer
On Error GoTo Err_IrParaControle
    NomeControle.SetFocus
Exit_IrParaControle:
    Exit Function
Err_IrParaControle:
    MsgBox Error$
    Resume Exit_IrParaControle
End Function

Function NullToZero(x As Variant) As Variant
On Error GoTo Err_NullToZero
    If IsNull(x) Then
        NullToZero = 0
    Else
        NullToZero = x
    End If
Exit_NullToZero:
    Exit Function
Err_NullToZero:
    MsgBox Error$
    Resume Exit_NullToZero
End Function

Function PLayWave(NomeArq As String) As Integer
On Error GoTo Err_PlayWave

    PLayWave = SndPlaySound(NomeArq, 1)
Exit_PlayWave:
    Exit Function
Err_PlayWave:
    MsgBox Error$
    Resume Exit_PlayWave
End Function

Function procurar(NomeForm As String, FindWhat As control, FindWhere As String, Find_A As Integer) As Integer
On Error GoTo Err_Procurar
Dim Invis�vel
    'DoCmd.DoMenuItem A_FORMBAR, A_FILE, A_SAVERECORD, , acMenuVer1X
    If IsNull(FindWhat) Then Exit Function
    DoCmd.OpenForm NomeForm
    'DoCmd.ShowAllRecords
    If Forms(NomeForm)(FindWhere).Visible = False Then
        Invis�vel = True
        Application.Echo False
        Forms(NomeForm)(FindWhere).Visible = True
    End If
        Forms(NomeForm)(FindWhere).SetFocus
        DoCmd.FindRecord FindWhat, Find_A
    If Invis�vel Then
        Application.Echo True
        SendKeysNovo ("{TAB}")
        Forms(NomeForm)(FindWhere).Visible = False
    End If
Exit_Procurar:
    Exit Function
Err_Procurar:
    MsgBox Error$
    Resume Exit_Procurar
End Function

Function Reconsultar()
On Error GoTo Err_Reconsultar
    Screen.ActiveForm.Requery
Exit_Reconsultar:
    Exit Function
Err_Reconsultar:
    MsgBox Error$
    Resume Exit_Reconsultar
End Function

Function ReconsultarControle(NomeControle As control) As Integer
On Error GoTo Err_ReconsultarControle
    NomeControle.Requery
Exit_ReconsultarControle:
    Exit Function
Err_ReconsultarControle:
    MsgBox Error$
    Resume Exit_ReconsultarControle
End Function
Function Sair()
On Error GoTo Err_Sair
    Application.Quit
Exit_Sair:
    Exit Function
Err_Sair:
    MsgBox Error$
    Resume Exit_Sair
End Function

Function Salvar()
On Error GoTo Err_Salvar
     DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70

Exit_Salvar:
    Exit Function
Err_Salvar:
    MsgBox Error$
    Resume Exit_Salvar
End Function

Function Senha()
On Error GoTo Err_Senha
Dim rs As Recordset, db As Database
Dim z, y
    z = Trim(Mid(str$(CLng(Date)), 1, 3) & Chr(Day(Date) / 2 + 65) & Mid(str$(CLng(Date)), 4, 4))
    y = InputBox("Digite a senha:")
    
    If y = z Then
        Senha = 1
        Set db = DBEngine.Workspaces(0).Databases(0)
        Set rs = db.OpenRecordset("tblAM_Acesso")
        rs.AddNew
        rs!DT_Acesso = Date
        rs.Update
        rs.Close
    Else
        Senha = 0
    End If
Exit_Senha:
    Exit Function
Err_Senha:
    MsgBox Error$
    Resume Exit_Senha
End Function

Function TBOff(TB As String) As Integer
On Error GoTo Err_TBOff
    DoCmd.ShowToolbar TB, A_TOOLBAR_NO
Exit_TBOff:
    Exit Function
Err_TBOff:
    MsgBox Error$
    Resume Exit_TBOff
End Function

Function TBOn(TB As String) As Integer
On Error GoTo Err_TBOn
    DoCmd.ShowToolbar TB, A_TOOLBAR_YES
Exit_TBOn:
    Exit Function
Err_TBOn:
    MsgBox Error$
    Resume Exit_TBOn
End Function

Function Zoom()
On Error GoTo Err_Zoom
    SendKeysNovo ("+{F2}")
Exit_Zoom:
    Exit Function
Err_Zoom:
    MsgBox Error$
    Resume Exit_Zoom
End Function

Function DVCGC(CGC As String)
On Error GoTo Err_CGC
Dim intSoma, intSoma1, intSoma2, intInteiro As Long
Dim intNumero, intMais, i, intResto As Integer
Dim intDig1, intDig2 As Integer
Dim strCampo, strCaracter, strConf, strCGC As String
Dim dblDivisao As Double
intSoma = 0
intSoma1 = 0
intSoma2 = 0
intNumero = 0
intMais = 0
'Separa os d�gitos do CGC que ser�o multiplicados de 2 a 9.
'Retira a "/" da m�scara de entrada.
strCGC = right(CGC, 6)
strCGC = left(strCGC, 4)
strCampo = left(CGC, 8)
strCampo = right(strCampo, 4) & strCGC
For i = 2 To 9
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma1 = intSoma1 + intMais
Next i
'Separa os 4 primeiros d�gitos do CGC
strCampo = left(CGC, 4)
For i = 2 To 5
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma2 = intSoma2 + intMais
Next i
intSoma = intSoma1 + intSoma2
dblDivisao = intSoma / 11
intInteiro = Int(dblDivisao) * 11
intResto = intSoma - intInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If
intSoma = 0
intSoma1 = 0
intSoma2 = 0
intNumero = 0
intMais = 0
strCGC = right(CGC, 6)
strCGC = left(strCGC, 4)
strCampo = left(CGC, 8)
strCampo = right(strCampo, 3) & strCGC & intDig1
For i = 2 To 9
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma1 = intSoma1 + intMais
Next i
strCampo = left(CGC, 5)
For i = 2 To 6
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    intSoma2 = intSoma2 + intMais
Next i
intSoma = intSoma1 + intSoma2
dblDivisao = intSoma / 11
intInteiro = Int(dblDivisao) * 11
intResto = intSoma - intInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CGC esteja errado dispara a mensagem
If strConf <> right(CGC, 2) Then
    MsgBox "O d�gito do CNPJ n�o est� correto.", 16, "Aten��o"
    DVCGC = False
    Exit Function
End If
    DVCGC = True
Exit Function
Exit_CGC:
    Exit Function
Err_CGC:
    MsgBox Error$
    Resume Exit_CGC
End Function

Function DVCPF(CPF As String)
On Error GoTo Err_CPF
Dim lngSoma, lngInteiro As Long
Dim intNumero, intMais, i, intResto As Integer
Dim intDig1, intDig2 As Integer
Dim strCampo, strCaracter, strConf As String
Dim dblDivisao As Double
lngSoma = 0
intNumero = 0
intMais = 0
strCampo = left(CPF, 9)
'Inicia c�lculos do 1� d�gito
'A fun��o Right() separa os caracteres da direita
'A fun��o Left() separa os caracteres da esquerda
'A fun��o Int() retorna o valor inteiro de um campo num�rico
For i = 2 To 10
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If
strCampo = strCampo & intDig1
lngSoma = 0
intNumero = 0
intMais = 0
For i = 2 To 11
    strCaracter = right(strCampo, i - 1)
    intNumero = left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> right(CPF, 2) Then
    MsgBox "O d�gito do CPF n�o est� correto.", 16, "Aten��o"
    DVCPF = False
    Exit Function
End If

DVCPF = True
Exit Function
Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function

Function CalculaData(data As Date) As Date
Dim sqlString As String
Dim rsFer As New ADODB.Recordset

Dim i As Integer

AbrirConexao

If Weekday(data) = 1 Then
  data = data + 1
End If
If Weekday(data) = 7 Then
  data = data + 2
End If

i = 0

For i = 0 To 7 Step 1
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(data, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    If rsFer.RecordCount = 0 Then
        rsFer.Close
        Exit For
    End If
    If rsFer.RecordCount <> 0 Then
    data = data + 1
      If Weekday(data) = 1 Then
        data = data + 1
      End If
      If Weekday(data) = 7 Then
        data = data + 2
      End If
    End If
    rsFer.Close
Next

CalculaData = data

End Function
Function CalculaDataCob(data As Date) As Date
Dim sqlString As String
Dim rsFer As New ADODB.Recordset
  
AbrirConexao

Dim i As Integer
'Dim db As DataBase
'Dim rs As Recordset
'Dim qry_Fer As QueryDef

If Weekday(data) = 1 Then
  data = data + 1
End If
If Weekday(data) = 7 Then
  data = data + 2
End If

i = 0

For i = 0 To 7 Step 1
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(data, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    'Set db = CurrentDb()
    'Set qry_Fer = db.QueryDefs("qryLctoFin_Feriados")
    'qry_Fer.Parameters(0) = Data
    'Set rs = qry_Fer.OpenRecordset()
    If rsFer.RecordCount = 0 Then
        rsFer.Close
        Exit For
    End If
    If rsFer.RecordCount <> 0 Then
    data = data + 1
      If Weekday(data) = 1 Then
        data = data + 1
      End If
      If Weekday(data) = 7 Then
        data = data + 2
      End If
    End If
    rsFer.Close
Next


'If Weekday(Data) + 1 = 7 Then
'    Data = Data + 3
'ElseIf Weekday(Data) + 1 = 1 Then
'    Data = Data + 3
'ElseIf Weekday(Data) + 1 = 2 Then
'    Data = Data + 2
'Else
'    Data = Data
'End If

i = 0

For i = 0 To 7 Step 1
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(data, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    
    If rsFer.RecordCount = 0 Then
        rsFer.Close
        Exit For
    End If
    
    If rsFer.RecordCount <> 0 Then
    data = data + 1
      If Weekday(data) = 1 Then
        data = data + 1
      End If
      If Weekday(data) = 7 Then
        data = data + 2
      End If
    End If
    rsFer.Close
Next
'Set db = Nothing


'SeImed(DiaSem([DTVcto_LctoFin]+1)=7;[DTVcto_LctoFin]+3;
'SeImed(DiaSem([DTVcto_LctoFin]+1)=1;[DTVcto_LctoFin]+3;
'SeImed(DiaSem([DTVcto_LctoFin]+1)=2;[DTVcto_LctoFin]+2;
'[DTVcto_LctoFin]+1)

CalculaDataCob = data

End Function


'Public Function GetUserLevel() As Long
'Dim gp As DAO.Group
'Dim lngCurLevel As Long
'Dim lngLevel As Long
'
'For Each gp In Workspaces(0).Users(CurrentUser()).Groups
'    Select Case gp.Name
'        Case "Full Permissions"
'            lngCurLevel = 3
'        Case "Fulll Data Users"
'            lngCurLevel = 2
'        Case "Users"
'            lngCurLevel = 1
'    End Select
'    lngLevel = IIf(lngLevel > lngCurLevel, lngLevel, lngCurLevel)
'Next gp
'
'GetUserLevel = lngLevel
'
'Set gp = Nothing
'
'End Function

Public Function GetUserGroup()

Dim wrkDefault As Workspace
    Dim usrNew As User
    Dim usrLoop As User
    Dim grpNew As Group
    Dim grpLoop As Group
    Dim grpMember As Group

    Set wrkDefault = DBEngine.Workspaces(0)

    With wrkDefault

        ' Cria e acrescenta o novo usu�rio.
        'Set usrNew = .CreateUser("Francisco Silva", _
        '    "abc123DEF456", "Senha1")
        '.Users.Append usrNew

        ' Cria e acrescenta o novo grupo.
        'Set grpNew = .CreateGroup("Contas", _
        '    "UVW987xyz654")

'.Groups.Append grpNew

        ' Torne o usu�rio Francisco Silva um membro do
        ' grupo Contas criando e adicionando o objeto
        ' Group adequado � cole��o Groups do usu�rio.
        ' O mesmo � conseguido se um objeto User
        ' que represente Francisco Silva for criado
        ' e acrescentado � cole��o Users do grupo
        ' Contas.
        'Set grpMember = usrNew.CreateGroup("Contas")
        'usrNew.Groups.Append grpMember

        'Debug.Print "Cole��o Users:"

        ' Enumera todos os objetos User na cole��o

' Users do espa�o de trabalho padr�o.
        For Each usrLoop In .Users
            Debug.Print "    " & usrLoop.Name
            Debug.Print "        Pertence a estes grupos:"

            ' Enumera todos os objetos Group em cada
            ' cole��o Groups do objeto User.
            If usrLoop.Groups.count <> 0 Then
                For Each grpLoop In usrLoop.Groups
                    Debug.Print "            " & _
                        grpLoop.Name
                Next grpLoop
            Else
                Debug.Print "            [Nenhum]"

End If

        Next usrLoop

        Debug.Print "Cole��o Groups:"

        ' Enumera todos os objetos Group na cole��o
        ' Groups do espa�o de trabalho padr�o.
        For Each grpLoop In .Groups
            Debug.Print "    " & grpLoop.Name
            Debug.Print "        Tem como seus membros:"

            ' Enumera todos os objetos User em cada
            ' cole��o Users do objeto Group.
            If grpLoop.Users.count <> 0 Then
                For Each usrLoop In grpLoop.Users
                    Debug.Print "            " & usrLoop.Name
                Next usrLoop
            Else
                Debug.Print "            [Nenhum]"
            End If

        Next grpLoop

        ' Exclui os objetos User e Group pois isto �
        ' somente uma demonstra��o.
        '.Users.Delete "Francisco Silva"
        '.Groups.Delete "Contas"

    End With

End Function


Public Function AbreSelArquivo(ByVal frmHwnd As Single, Optional Titulo As String = "Sele��o de arquivos", Optional DiretorioInicial As String = "C:\", Optional Filtro As String) As String
On Error GoTo Err_AbreSelArquivo
Dim hWnd
'API do Windows que mostra a janela de abertura de arquivo

' BACKUP DO PADR�O DE FILTRO ====>>>> "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    
    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    
    'Esta execu��o pode ser ignorada se der erro
    OFName.hwndOwner = hWnd
    
    
    If IsMissing(Filtro) Or Trim$(Filtro) = "" Then
     'Filtro de Arquivos
     OFName.lpstrFilter = "Imagens v�lidas" + Chr$(0) + "*.png;*.bmp;*.jpg;*.xls" + Chr$(0)

    Else
     OFName.lpstrFilter = Filtro
    End If

    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = DiretorioInicial
    OFName.lpstrTitle = Titulo
    
    OFName.Flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        AbreSelArquivo = Trim$(OFName.lpstrFile)
    Else
        AbreSelArquivo = ""
    End If
    
Exit_AbreSelArquivo:
    Exit Function
Err_AbreSelArquivo:
    MsgBox Err.Description
    Resume Exit_AbreSelArquivo
End Function


Public Function OpenGetFileDialog( _
    Optional ByRef DialogTitle As Variant, _
    Optional ByVal InitialDir As Variant, _
    Optional ByVal Filter As Variant _
) As String

On Error GoTo Err_OpenGetFileDialog

    Dim OpenFile    As OPENFILENAME
    Dim lReturn     As Long

    If IsMissing(DialogTitle) Then DialogTitle = "Default Title"
    If IsMissing(InitialDir) Then InitialDir = "C:\"
    If IsMissing(Filter) Then Filter = ""
    
    OpenFile.lStructSize = LenB(OpenFile)
    OpenFile.hwndOwner = Application.hWndAccessApp
    OpenFile.lpstrFile = String(256, 0)
    OpenFile.nMaxFile = LenB(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = InitialDir
    OpenFile.lpstrFilter = Filter
    OpenFile.lpstrTitle = DialogTitle
    OpenFile.Flags = 0
    
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        OpenGetFileDialog = ""
    Else
        OpenGetFileDialog = OpenFile.lpstrFile
    End If
Exit_OpenGetFileDialog:
    Exit Function
Err_OpenGetFileDialog:
    MsgBox Err.Description
    Resume Exit_OpenGetFileDialog
End Function
Function STRMaiuscula(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim x
  Dim nova As String
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "u"
      Case "�", "�"
        x = "c"
      Case "�", "�"
        x = "."
      Case "!", "@", "#", "%", "%", "^", "'", "&", "*", "_", "+", "=", ":", ";", "?", ">", "<", "~", "`", "|", "\", Chr$(34)
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STRMaiuscula = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function
Function STRMaiusculaSINT(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim x
  Dim nova As String
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�", "�", "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�", "�", "�", "�"
        x = "u"
      Case "�", "�"
        x = "c"
      Case "�", "�"
        x = ""
      Case "!", "@", "#", "%", "%", "^", "'", "&", "*", "_", "+", "=", ":", ";", "?", ">", "<", "~", "`", "|", "\", Chr$(34)
        x = " "
          Case ".", "-", "/", ",", "(", ")", "`", "~", "'", "�", "^"
        x = ""

      Case Else
      x = x
    End Select
    nova = nova & x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STRMaiusculaSINT = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function
Function STREspeciais(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim x
  Dim nova As String
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "�"
        x = "a"
      Case "�"
        x = "A"
      Case "�"
        x = "e"
      Case "�"
        x = "E"
      Case "�"
        x = "i"
      Case "�"
        x = "I"
      Case "�"
        x = "o"
      Case "�"
        x = "O"
      Case "�"
        x = "u"
      Case "�"
        x = "U"
      Case "�", "�"
        x = ""
      Case "^", "'", "~", "`", "|", """", "�", "!", "@", "#", "$", "%", "&"
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STREspeciais = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Function STRAcentos(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim x
  Dim nova As String
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�"
        x = "A"
      Case "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�"
        x = "E"
      Case "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�"
        x = "I"
      Case "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�"
        x = "O"
      Case "�", "�", "�", "�"
        x = "u"
      Case "�", "�", "�", "�"
        x = "U"
      Case "�"
        x = "c"
      Case "�"
        x = "C"
      Case "�", "�"
        x = "."
      Case "^", "'", "~", "`"
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STRAcentos = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function
Function STRUnicode(campo As Variant) As String
  On Error GoTo Err_STR
 
'� \u00e1
'� \u00e0
'� \u00e2
'� \u00e3
'� \u00e4
'� \u00c1
'� \u00c0
'� \u00c2
'� \u00c3
'� \u00c4
campo = Replace(campo, "\u00e1", "�")
campo = Replace(campo, "\u00e0", "�")
campo = Replace(campo, "\u00e2", "�")
campo = Replace(campo, "\u00e3", "�")
campo = Replace(campo, "\u00e4", "�")
campo = Replace(campo, "\u00c1", "�")
campo = Replace(campo, "\u00c0", "�")
campo = Replace(campo, "\u00c2", "�")
campo = Replace(campo, "\u00c3", "�")
campo = Replace(campo, "\u00c4", "�")
'� \u00e9
'� \u00e8
'� \u00ea
'� \u00c9
'� \u00c8
'� \u00ca
'� \u00cb
campo = Replace(campo, "\u00e9", "�")
campo = Replace(campo, "\u00e8", "�")
campo = Replace(campo, "\u00ea", "�")
campo = Replace(campo, "\u00c9", "�")
campo = Replace(campo, "\u00c8", "�")
campo = Replace(campo, "\u00ca", "�")
campo = Replace(campo, "\u00cb", "�")
'� \u00ed
'� \u00ec
'� \u00ee
'� \u00ef
'� \u00cd
'� \u00cc
'� \u00ce
'� \u00cf
campo = Replace(campo, "\u00ed", "�")
campo = Replace(campo, "\u00ec", "�")
campo = Replace(campo, "\u00ee", "�")
campo = Replace(campo, "\u00ef", "�")
campo = Replace(campo, "\u00cd", "�")
campo = Replace(campo, "\u00cc", "�")
campo = Replace(campo, "\u00ce", "�")
campo = Replace(campo, "\u00cf", "�")
'� \u00f3
'� \u00f2
'� \u00f4
'� \u00f5
'� \u00f6
'� \u00d3
'� \u00d2
'� \u00d4
'� \u00d5
'� \u00d6
campo = Replace(campo, "\u00f3", "�")
campo = Replace(campo, "\u00f2", "�")
campo = Replace(campo, "\u00f4", "�")
campo = Replace(campo, "\u00f5", "�")
campo = Replace(campo, "\u00f6", "�")
campo = Replace(campo, "\u00d3", "�")
campo = Replace(campo, "\u00d2", "�")
campo = Replace(campo, "\u00d4", "�")
campo = Replace(campo, "\u00d5", "�")
campo = Replace(campo, "\u00d6", "�")
'� \u00fa
'� \u00f9
'� \u00fb
'� \u00fc
'� \u00da
'� \u00d9
'� \u00db
campo = Replace(campo, "\u00fa", "�")
campo = Replace(campo, "\u00f9", "�")
campo = Replace(campo, "\u00fb", "�")
campo = Replace(campo, "\u00fc", "�")
campo = Replace(campo, "\u00da", "�")
campo = Replace(campo, "\u00d9", "�")
campo = Replace(campo, "\u00db", "�")
'� \u00e7
'� \u00c7
'� \u00f1
'� \u00d1
campo = Replace(campo, "\u00e7", "�")
campo = Replace(campo, "\u00c7", "�")
campo = Replace(campo, "\u00f1", "�")
campo = Replace(campo, "\u00d1", "�")

STRUnicode = campo
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function


Public Function MMCase(texto As String) As String
Dim sPalavra As String, iPosIni As Integer
Dim iPosFim As Integer, sResultado As String
iPosIni = 1
texto = LCase(texto) & " "

Do Until InStr(iPosIni, texto, " ") = 0
iPosFim = InStr(iPosIni, texto, " ")
sPalavra = Mid(texto, iPosIni, iPosFim - iPosIni)
iPosIni = iPosFim + 1
If sPalavra <> "de" And sPalavra <> "da" And _
sPalavra <> "do" And sPalavra <> "das" _
And sPalavra <> "dos" And sPalavra <> _
"a" And sPalavra <> "e" Then
sPalavra = UCase(left(sPalavra, 1)) & _
LCase(Mid(sPalavra, 2))
End If
sResultado = sResultado & " " & sPalavra
Loop
MMCase = Trim(sResultado)
End Function


Function TextoPor2(texto As String) As Variant

Dim arr(2) As Variant
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim parte1 As String
Dim parte2 As String
Dim parte3 As String

For i = 25 To 1 Step -1

If Mid(texto, i, 1) = " " Then
    Exit For
End If
    TextoPor2 = arr
Next
    parte1 = Mid(texto, 1, i)
    
For j = (50 - (25 - i)) To 1 Step -1
If Mid(texto, j, 1) = " " Then
    Exit For
End If
    TextoPor2 = arr
Next
    parte2 = Mid(texto, i + 1, j - i)
    
For k = 75 - (50 - j) To 1 Step -1
If Mid(texto, k, 1) = " " Then
    Exit For
End If
    TextoPor2 = arr
Next
    parte3 = Mid(texto, j + 1, 75)




arr(0) = parte1
arr(1) = parte2
arr(2) = parte3

TextoPor2 = arr

End Function

Public Sub ReIniciaMySqlUpdate()
'Limpa o Vetor de Campos para atualiza��o da base
Erase Campos
End Sub

Public Sub AlterarCampo(ByVal campo As ADODB.Field, Optional ByVal valor)
On Error GoTo NoBug
Dim PrecisaAspas As Boolean
Dim CId As Integer
Dim BoundErr As Boolean

'Verificando o tipo de campo do qual ser� trabalhado

'Se o valor for nulo, deixar como nulo
'Cada tipo de campo retorna um n�mero
If IsNull(valor) Then
 PrecisaAspas = False
 valor = Null
Else
 PrecisaAspas = False
 Select Case campo.Type
 Case 2    ' Tinyint
  PrecisaAspas = False
  valor = CInt(valor)
 Case 3    ' Integer
  PrecisaAspas = False
  valor = CInt(valor)
 Case 5    ' Double
  PrecisaAspas = False
  valor = CDbl(valor)
 Case 16   ' TinyInt
  PrecisaAspas = False
  valor = CInt(valor)
 Case 19
  PrecisaAspas = False
  valor = CInt(valor)
 Case 129  ' Char
  PrecisaAspas = True
  valor = CStr(valor)
 Case 133  ' DateTime
  PrecisaAspas = True
  valor = CStr(Format(valor, "yyyy/MM/dd"))
 Case 135  ' DateTime
  PrecisaAspas = True
  If Len(valor) <= 8 Then
    valor = CStr(Format(valor, "hh:mm:ss"))
  Else
    valor = CStr(Format(valor, "yyyy/MM/dd"))
  End If
 Case 200  ' VarChar
  valor = CStr(valor)
  PrecisaAspas = True
 Case 201  ' VarChar
  PrecisaAspas = True
  valor = CStr(valor)
 Case 202  ' VarChar
  PrecisaAspas = True
  valor = CStr(valor)
 Case 205 ' BLOB
  PrecisaAspas = False
  valor = ConverteParaBlob(CStr(valor))
 End Select
End If


'Se o tipo de campo solicitar as aspas (campo texto, data, blob, etc)
If PrecisaAspas = False And Not IsNull(valor) Then
 If Trim$(valor) = "" Then valor = 0
 valor = Replace(valor, ",", ".")
End If

BoundErr = True
CId = UBound(Campos) + 1
BoundErr = False

If IsNull(valor) Then
 valor = "NULL"
Else
 If PrecisaAspas Then valor = "'" & ConverteValor(CStr(valor)) & "'"
End If

'Criando mais um campo para altera��o
ReDim Preserve Campos(CId)
Campos(CId).campo = campo.Name
Campos(CId).valor = valor

Exit Sub

NoBug:
If BoundErr Then CId = 0: Resume Next
'MsgBox Err.Description
Resume Next
End Sub

Private Function ConverteParaBlob(ByVal valor As String) As String
'('), aspas duplas ("), barra invertida (\) e NUL (o byte NULL).
Dim imgValor As String
Dim n As Integer
'Muda um arquivo bin�rio (.gif, .jpg, .exe) para texto
n = FreeFile
Open valor For Binary Access Read As #n
 imgValor = Input(LOF(n), 1)
Close #n

'Converte caracteres que anulam a string
imgValor = Replace(imgValor, "'", "\'")
'imgValor = Replace(imgValor, "\", "\\")
imgValor = Replace(imgValor, vbCr, "\r")
imgValor = Replace(imgValor, vbLf, "\n")

'MsgBox imgValor
ConverteParaBlob = imgValor
End Function

Public Function ConverteValor(txt As String) As String
'Converte caracteres que anulam a string
'txt = Replace(txt, "\", "\\")
txt = Replace(txt, "'", "''")
'txt = Replace(txt, vbCr, "\r")
'txt = Replace(txt, vbLf, "\n")
ConverteValor = txt
End Function
Public Function GerarSQLUpdate(Optional ForAddNew As Boolean = False) As String
On Error GoTo NoBug
Dim i As Integer
Dim str As String
Dim A, b As String

'Gerando a string de INSERT OU UPDATE
'Se for INSERT, gerar no formato --> (Campo1, Campo2, Campo3...) VALUES (Valor1, Valor2, Valor3)
'Se for UPDATE, gerar no formato --> Campo1=Valor1, Campo2=Valor2, Campo3=Valor3...

If Not ForAddNew Then
 For i = 0 To UBound(Campos)
  str = str & Campos(i).campo & "=" & Campos(i).valor & ", "
 Next i
 str = Mid(str, 1, Len(str) - 2)
 GerarSQLUpdate = "SET " & str
Else
 For i = 0 To UBound(Campos)
  A = A & Campos(i).campo & ", "
  b = b & Campos(i).valor & ", "
 Next i
 A = Mid(A, 1, Len(A) - 2)
 b = Mid(b, 1, Len(b) - 2)
 GerarSQLUpdate = "(" & A & ") VALUES(" & b & ")"
End If

Exit Function

NoBug:
If Err.Number = 9 Then Exit Function

End Function

Public Sub SaveSQLString(strSQL As String)
On Error Resume Next
Dim i As String

i = FreeFile
Open "C:\sam\sql_string.txt" For Output As #i
 Print #i, strSQL
Close #i

Shell "notepad C:\sam\sql_string.txt"

End Sub

Function ExistePasta(path As String)

Dim drive As String
Dim pastas As String
Dim x

path = path & "\"

If InStr(path, "\") = 1 Then ' se for pasta pela rede

        'MsgBox "pasta na rede"
        drive = drive & left(path, InStr(3, path, "\") - 1)
        'MsgBox "Servidor �: " & drive & "\"
        x = CInt(Len(path)) - CInt(Len(drive))
        pastas = right(path, x)
        
        Call GerarPasta(drive, pastas)
Else

        'MsgBox " pasta local"
        drive = drive & left(path, InStr(2, path, "\") - 1)
        'MsgBox "Unidade �: " & drive & "\"
        x = CInt(Len(path)) - CInt(Len(drive))
        pastas = right(path, x)
                     
        Call GerarPasta(drive, pastas)

End If


End Function
Sub GerarPasta(sDrive As String, sDir As String)

Dim sBuild As String

While InStr(2, sDir, "\") > 1

    sBuild = sBuild & left(sDir, InStr(2, sDir, "\") - 1) & "\"
    sDir = Mid$(sDir, InStr(2, sDir, "\"))
    
    If Dir$(sDrive & sBuild, 16) = "" Then
        MkDir sDrive & sBuild
    End If
Wend
End Sub

Function STRNome(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim x
  Dim nova As String
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "@", "#", "&", "*", "_", ":", ";", "'", "|", "\", "/"
        x = " "
      Case Else
      x = x
    End Select
    nova = nova & x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STRNome = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Function fSearchFile(ByVal strFileName As String, _
            ByVal strSearchPath As String) As String
'Returns the first match found
    Dim lpBuffer As String
    Dim lngResult As Long
    fSearchFile = ""
    lpBuffer = String$(1024, 0)
    lngResult = apiSearchTreeForFile(strSearchPath, strFileName, lpBuffer)
    If lngResult <> 0 Then
        If InStr(lpBuffer, vbNullChar) > 0 Then
            fSearchFile = left$(lpBuffer, InStr(lpBuffer, vbNullChar) - 1)
        End If
    End If
End Function


Public Function TestarVinculosSQL() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQL = True

'Obtendo uma conex�o qualquer de uma tabela qualquer vinculada no sistema para testar o v�nculo
CurrentDb.QueryDefs.Delete "GetActualConnection"
ConnX = CurrentDb.CreateQueryDef("GetActualConnection", "SELECT Connect FROM MSysObjects WHERE Flags=537919488 AND Name like 'Cadastro de Produtos' AND NOT ISNULL(Connect);").OpenRecordset!Connect
CurrentDb.QueryDefs.Delete "GetActualConnection"

'Se a conex�o for satisfat�ria, sair da rotina
If Trim$(" " & ConnX) = "" Then TestarVinculosSQL = True: Exit Function
If TestarConexao(ConnX) Then Exit Function

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'SQL' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'SQL';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQL = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQL = False
Exit Function

NoBug:


If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function
Public Function TestarVinculosCEP() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosCEP = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'CEP' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'CEP';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosCEP = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosCEP = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function
Public Function TestarVinculosWH() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosWH = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'WH' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'WH';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosWH = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosWH = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function TestarVinculosTABSQL() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosTABSQL = True

'Obtendo uma conex�o qualquer de uma tabela qualquer vinculada no sistema para testar o v�nculo
CurrentDb.QueryDefs.Delete "GetActualConnection"
ConnX = CurrentDb.CreateQueryDef("GetActualConnection", "SELECT Connect FROM MSysObjects WHERE Flags=537919488 AND Name like 'Detalhe Produtos Vendidos_Antigo' AND NOT ISNULL(Connect);").OpenRecordset!Connect
CurrentDb.QueryDefs.Delete "GetActualConnection"

'Se a conex�o for satisfat�ria, sair da rotina
If Trim$(" " & ConnX) = "" Then TestarVinculosTABSQL = True: Exit Function
If TestarConexao(ConnX) Then Exit Function

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'TAB' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'TAB';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosTABSQL = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosTABSQL = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function
Public Function TestarVinculosSISSQL() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSISSQL = True

'Obtendo uma conex�o qualquer de uma tabela qualquer vinculada no sistema para testar o v�nculo
CurrentDb.QueryDefs.Delete "GetActualConnection"
ConnX = CurrentDb.CreateQueryDef("GetActualConnection", "SELECT Connect FROM MSysObjects WHERE Flags=537919488 AND Name like 'Cadastro de produtos_Sispedal' AND NOT ISNULL(Connect);").OpenRecordset!Connect
CurrentDb.QueryDefs.Delete "GetActualConnection"

'Se a conex�o for satisfat�ria, sair da rotina
If Trim$(" " & ConnX) = "" Then TestarVinculosSISSQL = True: Exit Function
If TestarConexao(ConnX) Then Exit Function

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'SIS' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'SIS';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSISSQL = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSISSQL = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function



Public Function TestarConexao(StringConexao As String) As Boolean
On Error GoTo NoBug
Dim ConnTest As New ADODB.Connection

'SEMPRE DEIXAR ESTE VALOR COMO TRUE
TestarConexao = True

'Abrindo a conex�o para verificar se est� correta.
ConnTest.Open StringConexao
ConnTest.Close

Exit Function
NoBug:
'Caso ocorra algum erro, a conex�o n�o � v�lida
TestarConexao = False
Exit Function
End Function

Function Trunca(dblNumero As Double, dblDecimais As Double) As Double
    dblNumero = dblNumero * 10 ^ dblDecimais
    If Mid(right(dblNumero, 2), 1, 1) = "," Or Mid(right(dblNumero, 3), 1, 1) = "," Or Mid(right(dblNumero, 4), 1, 1) = "," Then
        Trunca = Fix(dblNumero) / 10 ^ dblDecimais
    Else
        Trunca = (dblNumero) / 10 ^ dblDecimais
    End If
End Function
Function TruncaIPI(dblNumero As Double, dblDecimais As Double) As Double
    dblNumero = dblNumero * 10 ^ dblDecimais
    If Mid(right(dblNumero, 2), 1, 1) = "," Or Mid(right(dblNumero, 3), 1, 1) = "," Or Mid(right(dblNumero, 4), 1, 1) = "," Or Mid(right(dblNumero, 5), 1, 1) = "," Or Mid(right(dblNumero, 6), 1, 1) = "," Or Mid(right(dblNumero, 7), 1, 1) = "," Or Mid(right(dblNumero, 8), 1, 1) = "," Or Mid(right(dblNumero, 9), 1, 1) = "," Or Mid(right(dblNumero, 10), 1, 1) = "," Or Mid(right(dblNumero, 11), 1, 1) = "," Or Mid(right(dblNumero, 12), 1, 1) = "," Then
        TruncaIPI = Fix(dblNumero) / 10 ^ dblDecimais
    Else
        TruncaIPI = (dblNumero) / 10 ^ dblDecimais
    End If
End Function
Public Function Email_CDO(Remetente As String, Destinatario As String, _
assunto As String, Corpo As String, Optional CC As String, _
Optional BCC As String, Optional Anexo1 As String, Optional Anexo2 As String)
On Error GoTo Err_STR


Dim iMsg As Object
Dim iConf As Object
Dim strBody As String
Dim Flds As Variant
Dim rsConfig As New ADODB.Recordset
Dim str As String
            
Dim emailRemetente As String
Dim nomeRemetente As String
Dim emailBcc As String
Dim arquivos As String
Dim smtpCliente As String
Dim smtpPorta As String
Dim smtpSSL As String
Dim smtpUsuario As String
Dim smtpSenha As String

AbrirConexao

If txtEnvioBol = 1 Then
    emailRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolUsuario'")
    nomeRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolNomeEmail'")
    'emailDestinatario = rsCob!Email
    emailBcc = BCC
    assunto = assunto
    smtpCliente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSMTP'")
    smtpPorta = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolPorta'")
    smtpSSL = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSSL'")
    smtpUsuario = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolUsuario'")
    smtpSenha = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSenha'")
    
ElseIf txtEnvioBol = 2 Then
    emailRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecEmail'")
    nomeRemetente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecNomeEmail'")
    'emailDestinatario = rsCob!Email
    emailBcc = ""
    assunto = assunto
    smtpCliente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecSMTP'")
    smtpPorta = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecPorta'")
    smtpSSL = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecSSL'")
    smtpUsuario = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecUsuario'")
    smtpSenha = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecSenha'")
ElseIf txtEnvioBol = 3 Then
    emailRemetente = DLookup("[EmailUser]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")
    nomeRemetente = Remetente
    emailBcc = ""
    assunto = assunto
    smtpCliente = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSMTP'")
    smtpPorta = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolPorta'")
    smtpSSL = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolSSL'")
    smtpUsuario = DLookup("[EmailUser]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")
    smtpSenha = DLookup("[EmailSenha]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")

End If

 
If emailRemetente = "" Then
    MsgBox "E-mail n�o enviado, verifique configura��es do Remetente e Destinat�rio", vbInformation, "Aten��o"
    Exit Function
End If

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
 
        iConf.Load -1    ' CDO Source Defaults
        Set Flds = iConf.Fields
        With Flds
            .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPorta
            .item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpCliente
            .item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
            .item("http://schemas.microsoft.com/cdo/configuration/sendusername") = smtpUsuario
            .item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtpSenha
            .item("http://schemas.microsoft.com/cdo/configuration/sendemailaddress") = emailRemetente
            .Update
        End With
 
    strBody = Corpo
    With iMsg
        Set .Configuration = iConf
        .To = Destinatario
        .CC = CC
        If txtEnvioBol = 1 Then
            .ReplyTo = DLookup("[Val_Par]", "tblParametro", "[Descr_Par] = 'CRecBolUsuario'")
            .BCC = emailBcc
        ElseIf txtEnvioBol = 2 Then
            .ReplyTo = "financeiro@proparts.esp.br" & IIf(BCC = "", "", ";" & BCC)
        ElseIf txtEnvioBol = 3 Then
            .ReplyTo = DLookup("[EmailUser]", "Vendedores", "[C�digoDoFuncion�rio] = '" & BCC & "'")
            .BCC = emailBcc
        End If
        .FROM = "" & nomeRemetente & " <" & emailRemetente & ">"
        .Subject = assunto
        .TextBody = strBody
        If Anexo1 <> "" Then
            .Addattachment Anexo1
        End If
        If Anexo2 <> "" Then
            .Addattachment Anexo2
        End If
        .send
    End With
 
    Set iMsg = Nothing
    Set iConf = Nothing

Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR

End Function

Function STRArroba(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim nova As Double
  Dim x
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "@"
        x = 1
      Case Else
      x = 0
    End Select
    nova = nova + x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STRArroba = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Function Est�Aberto(NomeFormul�rio As String) As Integer
On Error GoTo Err_Est�Aberto
Dim x As Integer
Dim N�meroDeFormul�rios
    N�meroDeFormul�rios = Forms.count                                            ' formul�rios.
    For x = 0 To N�meroDeFormul�rios - 1
        If Forms(x).Name = NomeFormul�rio Then
            Est�Aberto = -1
            Exit Function
        Else
            Est�Aberto = 0
        End If
    Next x
Exit_Est�Aberto:
    Exit Function
Err_Est�Aberto:
    MsgBox Error$
    Resume Exit_Est�Aberto
End Function

Public Function FSelDir(Optional InitialDir As String = "C:\") As String
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo

'API do Windows que exibe a janela de sele��o de diret�rios

With udtBI
       '.lpszTitle = lstrcat("C:\a&m", "Teste")
       .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = left$(sPath, iNull - 1)
        End If
    End If

    FSelDir = Trim$(" " & Replace(sPath, Chr(0), ""))
End Function

Function SendKeysNovo(txtTecla) As String
Dim ws As Object
Set ws = CreateObject("Wscript.shell")
ws.SendKeys txtTecla
Set ws = Nothing
End Function

Function STRTiraAcentos(campo As Variant) As String
  On Error GoTo Err_STR
  Dim A As Integer
  Dim x
  Dim nova As String
  A = 1
  x = Mid(campo, A, 1)
  While (A <= Len(campo))
    Select Case x
      Case "�", "�", "�", "�", "�"
        x = "a"
      Case "�", "�", "�", "�", "�"
        x = "A"
      Case "�", "�", "�", "�"
        x = "e"
      Case "�", "�", "�", "�"
        x = "E"
      Case "�", "�", "�", "�"
        x = "i"
      Case "�", "�", "�", "�"
        x = "I"
      Case "�", "�", "�", "�", "�"
        x = "o"
      Case "�", "�", "�", "�", "�"
        x = "O"
      Case "�", "�", "�", "�"
        x = "u"
      Case "�", "�", "�", "�"
        x = "U"
      Case "�"
        x = "c"
      Case "�"
        x = "C"
      Case "!", "@", "#", "%", "^", "&", "_", "~", "`", "\", "�", Chr$(34), "�", "�"
      x = " "
      Case Else
      x = (x)
    End Select
    nova = nova & x
    A = A + 1
    If (A <= Len(campo)) Then
      x = Mid(campo, A, 1)
    End If
  Wend
  STRTiraAcentos = nova
Exit_STR:
    Exit Function
Err_STR:
  MsgBox Error$
  Resume Exit_STR
End Function

Public Sub RegistrarErro(Optional Numero As Double, Optional Descricao As String, Optional Modulo As String, Optional Funcao As String)
Dim MyPath As String
Dim parts() As String
Dim i As Integer
Dim ErrorFile As String
Dim FN As Integer
Dim strError As String

'Gera o arquivo de log de erro em A&M_SIS_Diretorio\pml_err.log

strError = "Log de erro gerado em : " & Format(Now, "dd/mm/yyyy") & " �s " & Format(Now, "hh:mm:ss") & vbCrLf
strError = strError & "M�dulo : " & Modulo & vbCrLf & "Fun��o : " & Funcao & vbCrLf
strError = strError & "N�mero : " & Numero & vbCrLf & "Descri��o : " & Descricao & vbCrLf & vbCrLf

MyPath = CurrentDb.Name

parts = Split(MyPath, "\")

ErrorFile = ""
For i = 0 To UBound(parts) - 1
 ErrorFile = ErrorFile & parts(i) & "\"
Next i
ErrorFile = ErrorFile & "Sisparts_err.log"

FN = FreeFile
Open ErrorFile For Append As #FN
 Print #FN, strError
Close #FN


End Sub

Function FindAndReplace(ByVal strInString As String, _
        strFindString As String, _
        strReplaceString As String) As String
Dim intPtr As Integer
    If Len(strFindString) > 0 Then  'catch if try to find empty string
        Do
            intPtr = InStr(strInString, strFindString)
            If intPtr > 0 Then
                FindAndReplace = FindAndReplace & left(strInString, intPtr - 1) & _
                                        strReplaceString
                    strInString = Mid(strInString, intPtr + Len(strFindString))
            End If
        Loop While intPtr > 0
    End If
    FindAndReplace = FindAndReplace & strInString
End Function

Function Fix6(x As Variant) As Currency
On Error GoTo Err_Fix6
    Fix6 = Int(Nz(x) * 1000000 + 0.5) / 1000000
Exit_Fix6:
    Exit Function
Err_Fix6:
    MsgBox Error$
    Resume Exit_Fix6
End Function

Public Function TestarVinculosSQLSP() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQLSP = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'NFESP' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'NFESP';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQLSP = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQLSP = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function TestarVinculosSQLES() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQLES = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'NFEES' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'NFEES';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQLES = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQLES = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function TestarVinculosSQLSC() As Boolean
'Esta rotina verifica se os v�nculos das tabelas est�o corretos
On Error GoTo NoBug
Dim ConnX As String
Dim RecX As DAO.Recordset
Dim QDef As QueryDef
Dim TDef As TableDef

TestarVinculosSQLSC = True

'Obtendo TODAS as conex�es para obter uma v�lida
CurrentDb.QueryDefs.Delete "GetAllConnections"
Set QDef = CurrentDb.CreateQueryDef("GetAllConnections", "SELECT String_Con AS Conexao, ID_Con AS Id, Nome_Con AS Nome FROM tblConexao WHERE Tipo_Con = 'NFESC' ORDER BY FlagPadrao_Con ASC;")
Set RecX = QDef.OpenRecordset
CurrentDb.QueryDefs.Delete "GetAllConnections"

QDef.Close
Set QDef = Nothing

'Abrindo no sistema TODAS as tabelas vinculadas para atualizar os v�nculos
Do While Not RecX.EOF
 
 If TestarConexao(RecX!Conexao) Then
  
  'Setando a conex�o v�lida como conex�o padr�o, limpando os flags primeiro
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=0 WHERE Tipo_Con = 'NFESC';"
  CurrentDb.Execute "UPDATE tblConexao SET FlagPadrao_Con=-1 WHERE ID_Con=" & RecX!ID & ";"
  
  For Each TDef In CurrentDb.TableDefs
   If Trim$(" " & TDef.Connect) <> "" Then
    'Atualizando o v�nculo
    TDef.Connect = "ODBC;" & RecX!Conexao
    TDef.RefreshLink
   End If
  Next
  
  'MsgBox "V�nculo do banco de dados trocado para " & RecX!Nome & ".", vbOKOnly + vbInformation, "Aten��o"
  TestarVinculosSQLSC = True
  RecX.Close
  
  Set RecX = Nothing
  Exit Function
  
 End If
 RecX.MoveNext
Loop

RecX.Close
Set RecX = Nothing
TestarVinculosSQLSC = False
Exit Function

NoBug:

If Err.Number = 3265 Then Resume Next
'MsgBox Err.Number & Err.Description
Resume Next

End Function

Public Function UserAltStatusCob(txtID As Long, txtAlt As String, txtIDStatusCob As Long, txtAberto As String)
Dim txtDescrTPStatusCob As String
Dim txtMsg As String
AbrirConexao
If IsNull(txtAlt) Or txtAlt = "" Then
    txtDescrTPStatusCob = DLookup("[Descr_TPStatus]", "tblTPStatus", "[ID_TPStatus] = " & txtIDStatusCob & "")
    txtMsg = Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm") & " " & txtDescrTPStatusCob
    CNN.Execute "UPDATE tblLctoFin SET " _
    & "tblLctoFin.UserAlt_LctoFin = '" & txtMsg & "' " _
    & "WHERE (((tblLctoFin.ID_LctoFin)= " & txtID & "));", dbSeeChanges
Else
    txtDescrTPStatusCob = DLookup("[Descr_TPStatus]", "tblTPStatus", "[ID_TPStatus] = " & txtIDStatusCob & "")
    txtMsg = txtAlt & "->" & Format(Date, "dd/mm/yyyy") & " " & Format(Time(), "hh:mm") & " " & txtDescrTPStatusCob
    CNN.Execute "UPDATE tblLctoFin SET " _
    & "tblLctoFin.UserAlt_LctoFin = '" & txtMsg & "' " _
    & "WHERE (((tblLctoFin.ID_LctoFin)= " & txtID & "));", dbSeeChanges
End If
If txtAberto = "S" Then
    Forms!frmLctoFin_CntRec_Cadastro!UserAlt_LctoFin = txtMsg
End If
End Function

Public Function fncDiasUteis(dataLan�amento As Date, DataRef As Date) As Integer
Dim j%, dataAnalisada As Date
Dim sqlString As String
Dim rsFer As New ADODB.Recordset
Dim txtFeriado As Boolean

dataAnalisada = dataLan�amento '+ 1
AbrirConexao
Do While Not dataAnalisada > DataRef
    txtFeriado = False
    sqlString = "SELECT tblFeriados.Dt_Fer FROM tblFeriados WHERE (((tblFeriados.Dt_Fer)='" & Format(dataAnalisada, "yyyy/mm/dd") & "'));"
    rsFer.CursorLocation = adUseClient
    rsFer.CursorType = adOpenKeyset
    rsFer.LockType = adLockOptimistic
    rsFer.Open sqlString, CNN
    If rsFer.RecordCount = 0 Then
        rsFer.Close
    Else
        txtFeriado = True
        rsFer.Close
    End If

    If Eval("weekday(#" & Format(dataAnalisada, "mm/dd/yyyy") & "#) between 2 and 6") And txtFeriado = False Then
            j = j + 1
    End If
    dataAnalisada = dataAnalisada + 1
Loop
fncDiasUteis = j
End Function


>>>>>>> f4084cb29d769387d25e7b837853d2119e0da429
>>>>>>> ca95ee3e8bcb0745be1525054e4155ff5a288f06:referencia/modFuncoesGerais.bas
