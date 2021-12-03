Attribute VB_Name = "Module1"
Option Compare Database

Sub CriarTabelaEmBancoParaExportacao()
Dim tdfNew As DAO.TableDef: Set tdfNew = Application.CurrentDb.CreateTableDef("tmp__AILTON")
    
    tdfNew.Fields.Append tdfNew.CreateField("NOME", dbText)
    Application.CurrentDb.TableDefs.Append tdfNew

End Sub
