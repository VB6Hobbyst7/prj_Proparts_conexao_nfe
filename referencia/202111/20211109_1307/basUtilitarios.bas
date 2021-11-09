Attribute VB_Name = "basUtilitarios"
Option Compare Database
Option Explicit

Sub InformaErro(procName As String)
    MsgBox "Erro n� " & Err.Number & " @@" & Err.Description, _
        vbExclamation, "Procedimento: " & procName
End Sub

Function IsLoaded(ByVal strFormName As String) As Integer
 ' Retorna True se o formul�rio especificado estiver
 ' aberto no modo Formul�rio ou no modo Folha de Dados.
    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
End Function

