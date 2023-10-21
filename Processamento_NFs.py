import pyodbc
import xml.etree.ElementTree as ET
import os

# Conecte-se ao banco de dados Access
conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\murilo.barbosa\OneDrive - Corttex Industria Textil Ltda\Logistica\Logistica.accdb;')

# Crie um cursor para executar consultas SQL
cursor = conn.cursor()

# Pasta contendo os arquivos XML
pasta_xml = r'C:\Users\murilo.barbosa\OneDrive - Corttex Industria Textil Ltda\Bases\xml\Nova pasta'

# Loop pelos arquivos XML na pasta
for arquivo_xml in os.listdir(pasta_xml):
    if arquivo_xml.endswith('.xml'):
        arquivo_xml = os.path.join(pasta_xml, arquivo_xml)

        # Analise o arquivo XML
        tree = ET.parse(arquivo_xml)
        root = tree.getroot()

        # Acesse os elementos desejados usando namespaces
        namespace = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        natOp = root.find(".//nfe:ide/nfe:natOp", namespaces=namespace).text
        nNF = root.find(".//nfe:ide/nfe:nNF", namespaces=namespace).text
        dhEmi = root.find(".//nfe:ide/nfe:dhEmi", namespaces=namespace).text

        # Acesse o elemento "emit" e, em seguida, os elementos "xNome" e "xFant" dentro dele
        emit = root.find(".//nfe:emit", namespaces=namespace)
        cnpj = emit.find(".//nfe:CNPJ", namespaces=namespace).text
        xNomeEmit = emit.find(".//nfe:xNome", namespaces=namespace).text
        xFant = emit.find(".//nfe:xFant", namespaces=namespace).text

        # Acesse o elemento "dest" e, em seguida, os elementos desejados dentro dele
        dest = root.find(".//nfe:dest", namespaces=namespace)
        
        # Verifique se cada elemento existe antes de acessar sua propriedade "text"
        xNomeDest = dest.find(".//nfe:xNome", namespaces=namespace).text if dest.find(".//nfe:xNome", namespaces=namespace) is not None else ""
        xLgr = dest.find(".//nfe:xLgr", namespaces=namespace).text if dest.find(".//nfe:xLgr", namespaces=namespace) is not None else ""
        nro = dest.find(".//nfe:nro", namespaces=namespace).text if dest.find(".//nfe:nro", namespaces=namespace) is not None else ""
        xCpl = dest.find(".//nfe:xCpl", namespaces=namespace).text if dest.find(".//nfe:xCpl", namespaces=namespace) is not None else ""
        xBairro = dest.find(".//nfe:xBairro", namespaces=namespace).text if dest.find(".//nfe:xBairro", namespaces=namespace) is not None else ""
        xMun = dest.find(".//nfe:xMun", namespaces=namespace).text if dest.find(".//nfe:xMun", namespaces=namespace) is not None else ""
        UF = dest.find(".//nfe:UF", namespaces=namespace).text if dest.find(".//nfe:UF", namespaces=namespace) is not None else ""
        CEP = dest.find(".//nfe:CEP", namespaces=namespace).text if dest.find(".//nfe:CEP", namespaces=namespace) is not None else ""
        xPais = dest.find(".//nfe:xPais", namespaces=namespace).text if dest.find(".//nfe:xPais", namespaces=namespace) is not None else ""

        # Acesse o elemento "total" e, em seguida, os elementos "vFrete" e "vNF" dentro dele
        total = root.find(".//nfe:total", namespaces=namespace)
        vFrete = total.find(".//nfe:vFrete", namespaces=namespace).text if total.find(".//nfe:vFrete", namespaces=namespace) is not None else ""
        vNF = total.find(".//nfe:vNF", namespaces=namespace).text if total.find(".//nfe:vNF", namespaces=namespace) is not None else ""

        transp = root.find(".//nfe:transp", namespaces=namespace)

        # Extrair dados de <qVol>, <esp>, <nVol>, <pesoL> e <pesoB>
        CNPJTransp = transp.find(".//nfe:CNPJ", namespaces=namespace).text if transp.find(".//nfe:CNPJ", namespaces=namespace) is not None else ""
        xNomeTransp = transp.find(".//nfe:xNome", namespaces=namespace).text if transp.find(".//nfe:xNome", namespaces=namespace) is not None else ""
        IETransp = transp.find(".//nfe:IE", namespaces=namespace).text if transp.find(".//nfe:IE", namespaces=namespace) is not None else ""
        
        # Extrair dados de <qVol>, <esp>, <nVol>, <pesoL> e <pesoB>
        qVol = transp.find(".//nfe:qVol", namespaces=namespace).text if transp.find(".//nfe:qVol", namespaces=namespace) is not None else ""
        esp = transp.find(".//nfe:esp", namespaces=namespace).text if transp.find(".//nfe:esp", namespaces=namespace) is not None else ""
        nVol = transp.find(".//nfe:nVol", namespaces=namespace).text if transp.find(".//nfe:nVol", namespaces=namespace) is not None else ""
        pesoL = transp.find(".//nfe:pesoL", namespaces=namespace).text if transp.find(".//nfe:pesoL", namespaces=namespace) is not None else ""
        pesoB = transp.find(".//nfe:pesoB", namespaces=namespace).text if transp.find(".//nfe:pesoB", namespaces=namespace) is not None else ""

        # Execute a inserção SQL
        cursor.execute("INSERT INTO tb_NFs (natOp, nNF, dhEmi, CNPJ, xNomeEmit, xFant, xNomeDest, xLgr, nro, xCpl, xBairro, xMun, UF, CEP, xPais, vFrete, vNF, qVol, esp, nVol, pesoL, pesoB, CNPJTransp, xNomeTransp, IETransp) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (natOp, nNF, dhEmi, cnpj, xNomeEmit, xFant, xNomeDest, xLgr, nro, xCpl, xBairro, xMun, UF, CEP, xPais, vFrete, vNF, qVol, esp, nVol, pesoL, pesoB, CNPJTransp, xNomeTransp, IETransp))


# Confirme as alterações e feche a conexão
conn.commit()
conn.close()

print("Processo finalizado! Notas fiscais cadastradas no banco!")