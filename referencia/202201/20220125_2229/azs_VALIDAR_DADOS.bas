Attribute VB_Name = "azs_VALIDAR_DADOS"
Option Compare Database


''' #20211128_MoverArquivosProcessados
'Sub MoverArquivosProcessados()
'On Error GoTo adm_Err
'Dim db As DAO.Database: Set db = CurrentDb
'Dim rst As DAO.Recordset
'
'''#registroProcessado
''' 0|1 - caminhoDeColeta
''' 2 - caminhoDeColeta
''' 3 - caminhoDeColetaProcessados
''' 4 - caminhoDeColetaExpurgo
''' 5 - caminhoDeColetaAcoes
''' 8 - MOVER ARQUIVOS            -- CONTROLE INTERNO
''' 9 - FINAL                     -- CONTROLE INTERNO
'
''Dim sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar As String: sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar = _
''    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 1;"
''    Application.CurrentDb.Execute Replace(sql_Update_tblDadosConexaoNFeCTe_registroProcessado_Reclassificar, 1, 5)
'
'
''' Processados - registroProcessado(3)
''' 1. Classificar como "registroProcessado(3) - Processados" onde ...
''' 1.1 Onde temos arquivo para processar - "CaminhoDoArquivo" e
''' 1.2 Registro é valido "registroValido(1) - Registro Válido" e
''' 1.3 Registro processado é "registroProcessado (2) - CadastroDeComprasEmServidor()".
'Dim sql_Update_registroProcessado_Processados As String: sql_Update_registroProcessado_Processados = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 3 WHERE ((Not (tblDadosConexaoNFeCTe.CaminhoDoArquivo) Is Null) AND ((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=2));"
'    Application.CurrentDb.Execute sql_Update_registroProcessado_Processados
'
''' Expurgo - registroProcessado(4)
''' 1. Classificar como "registroProcessado(4) - Expurgo" onde ...
''' 1.1 Registro é valido "registroValido(1) - Registro Válido" e não foi processado "registroProcessado(0) - Não foi processado"
''' 2 Registro é invalido "registroValido(0) - Registro Inválido"
'Dim sql_Update_registroProcessado_Expurgo() As Variant: sql_Update_registroProcessado_Expurgo = Array( _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=1) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));", _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 4 WHERE (((tblDadosConexaoNFeCTe.registroValido)=0) AND ((tblDadosConexaoNFeCTe.registroProcessado)=0));")
'    executarComandos sql_Update_registroProcessado_Expurgo
'
''' Finalizados - registroProcessado(9)
'Dim sql_Update_CopyFile_Final As String: sql_Update_CopyFile_Final = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.registroProcessado = 9, tblDadosConexaoNFeCTe.CaminhoDoArquivo = [tblDadosConexaoNFeCTe].[CaminhoDestino], tblDadosConexaoNFeCTe.CaminhoDestino = Null WHERE (((tblDadosConexaoNFeCTe.[registroProcessado])=8))"
'
''' Atualização do caminho de destino
'Dim sql_Update_CaminhoDestino As String: sql_Update_CaminhoDestino = _
'    "UPDATE tblDadosConexaoNFeCTe SET tblDadosConexaoNFeCTe.CaminhoDestino = strCaminhoDestino([tblDadosConexaoNFeCTe].[CaminhoDoArquivo]), tblDadosConexaoNFeCTe.registroProcessado = 8 WHERE (((tblDadosConexaoNFeCTe.registroProcessado)<8));"
'    Application.CurrentDb.Execute sql_Update_CaminhoDestino
'
''' Seleção de arquivos para movimentação de pastas
'Dim sql_Select_CaminhoDestino As String: sql_Select_CaminhoDestino = _
'    "SELECT tblDadosConexaoNFeCTe.CaminhoDoArquivo, tblDadosConexaoNFeCTe.CaminhoDestino FROM tblDadosConexaoNFeCTe WHERE (((tblDadosConexaoNFeCTe.registroProcessado)=8));"
'
'    Debug.Print "############################"
'
'    '' MOVER ARQUIVOS
'    Set rst = db.OpenRecordset(sql_Select_CaminhoDestino)
'    Do While Not rst.EOF
'        If (Dir(rst.Fields("CaminhoDoArquivo").value) <> "") Then
'            FileCopy rst.Fields("CaminhoDoArquivo").value, rst.Fields("CaminhoDestino").value
'            Kill rst.Fields("CaminhoDoArquivo").value
'        End If
'        rst.MoveNext
'    Loop
'
'    '' ATUALIZAÇÃO - registroProcessado ( FINAL )
'    Application.CurrentDb.Execute sql_Update_CopyFile_Final
'
'db.Close
'
'
'adm_Exit:
'    Set db = Nothing
'    Set rst = Nothing
'
'    Exit Sub
'
'adm_Err:
'    Debug.Print "MoverArquivosProcessados() - " & Err.Description
'    TextFile_Append CurrentProject.path & "\" & strLog(), Err.Description
'    Resume adm_Exit
'
'End Sub
'
'
