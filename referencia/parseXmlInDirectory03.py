import os
from xml.dom.minidom import parse, parseString

listDirectory = ["C:\\temp\\68.365.5010002-96 - Proparts Comércio de Artigos Esportivos e Tecnologia Ltda\\recebimento\\"]
for directory in (listDirectory):
    for filename in os.listdir(directory):
        if filename.endswith(".xml"):
            with open(os.path.join(directory, filename)) as xmldata:
                try:

                    xml = parseString(xmldata.read())
                    xml_pretty_str = xml.toprettyxml()
                    nomaArquivo = filename.split(".")[0] + ".txt"
                    f = open(os.path.join(directory, nomaArquivo), "a+")
                    f.write(xml_pretty_str)
                    f.close

                    print(filename.split(".")[0])

                except OSError as e:
                    print("#### " + e.filename)