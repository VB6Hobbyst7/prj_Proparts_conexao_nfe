Attribute VB_Name = "azsCadastroRegistros"
Option Compare Database

Sub start_cadastro()

CadastroDeRegistros registro

End Sub


Private Sub CadastroDeRegistros(Itens As Collection)
Dim con As ADODB.Connection: Set con = CurrentProject.Connection
Dim i As Variant

    For Each i In Itens
        con.Execute i
    Next i

Set con = Nothing

End Sub


Private Function registro() As Collection
Set registro = New Collection

    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS40/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS41/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS50/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS60/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/CSOSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN102/CSOSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN500/CSOSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS40/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS41/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS50/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS60/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN102/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN500/orig',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/vCredICMSSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/modBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/modBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pICMS',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMSSN101/pCredSN',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pMVAST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pRedBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/pRedBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS00/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS20/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS51/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vBCST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS10/vICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS30/vICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS70/vICMSST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/ICMS|ICMS90/vICMSST',1)"
    
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI|cEnq',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|vBC',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|CST',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|pIPI',1)"
    registro.add "INSERT INTO tblOrigemDestino (Destino,Tag,TagOrigem) VALUES('tblCompraNFItem','nfeProc/NFe/infNFe/det|imposto/IPI/IPITrib|vIPI',1)"
End Function
