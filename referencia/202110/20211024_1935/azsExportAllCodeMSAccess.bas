Attribute VB_Name = "azsExportAllCodeMSAccess"

'' 02. EXPORT CODE
Public Sub ExportAllCode()
'' https://stackoverflow.com/questions/16948215/exporting-ms-access-forms-and-class-modules-recursively-to-text-files

    Dim C As VBComponent
    Dim Sfx As String
    Dim pathExit As String: pathExit = Replace(CurrentProject.path, left(CurrentProject.Name, Len(CurrentProject.Name) - 6), "") & "referencia\" & strControle & "\"
    
    For Each C In Application.VBE.VBProjects(1).VBComponents
        Select Case C.Type
            Case vbext_ct_ClassModule, vbext_ct_Document
                Sfx = ".cls"
            Case vbext_ct_MSForm
                Sfx = ".frm"
            Case vbext_ct_StdModule
                Sfx = ".bas"
            Case Else
                Sfx = ""
        End Select

        If Sfx <> "" Then
            CreateDir pathExit
            C.Export fileName:=pathExit & C.Name & Sfx
        End If
    Next C


Shell "explorer " & pathExit, vbMaximizedFocus

End Sub

'' Criar estrutura de diretorios
Private Function CreateDir(strPath As String)
    Dim elm As Variant
    Dim strCheckPath As String

    strCheckPath = ""
    For Each elm In Split(strPath, "\")
        strCheckPath = strCheckPath & elm & "\"
        If Len(Dir(strCheckPath, vbDirectory)) = 0 Then MkDir strCheckPath
    Next
End Function

'' 01. Add VBIDE (Microsoft Visual Basic for Applications Extensibility 5.3
Private Sub AddRefGuid()
On Error Resume Next

    Application.VBE.VBProjects(1).References.AddFromGuid _
        "{0002E157-0000-0000-C000-000000000046}", 2, 0
 
End Sub

