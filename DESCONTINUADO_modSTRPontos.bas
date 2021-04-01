Attribute VB_Name = "modSTRPontos"
Option Compare Database

'Function STRPontos(campo As Variant) As String
'  On Error GoTo Err_STR
'  Dim a As Integer
'  Dim nova As String
'  Dim x
'  a = 1
'  x = Mid(campo, a, 1)
'  While (a <= Len(campo))
'    Select Case x
'      Case ".", ",", "-", " ", "/", "\"
'        x = ""
'      Case Else
'      x = UCase$(x)
'    End Select
'    nova = nova & x
'    a = a + 1
'    If (a <= Len(campo)) Then
'      x = Mid(campo, a, 1)
'    End If
'  Wend
'  STRPontos = nova
'Exit_STR:
'    Exit Function
'Err_STR:
'  MsgBox Error$
'  Resume Exit_STR
'End Function
