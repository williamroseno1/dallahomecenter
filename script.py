import pandas as pd
from pathlib import Path
import xml.etree.ElementTree as ET

arquivo = Path("data/ds460.xls")


def ler_planilha_xml(caminho):
    namespace = "{urn:schemas-microsoft-com:office:spreadsheet}"
    raiz = ET.parse(caminho).getroot()
    tabela = raiz.find(f".//{namespace}Table")
    linhas = []

    for linha_xml in tabela.findall(f"{namespace}Row"):
        valores = []
        for celula in linha_xml.findall(f"{namespace}Cell"):
            indice = celula.get(f"{namespace}Index")
            if indice is not None:
                valores.extend([None] * (int(indice) - 1 - len(valores)))

            dado = celula.find(f"{namespace}Data")
            valor = dado.text if dado is not None else None
            tipo = dado.get(f"{namespace}Type") if dado is not None else None

            if valor is not None and tipo == "Number":
                numero = float(valor)
                valor = int(numero) if numero.is_integer() else numero
            elif valor is not None and tipo == "Boolean":
                valor = valor == "1"
            elif valor is not None and tipo == "DateTime":
                valor = pd.to_datetime(valor)

            valores.append(valor)
        linhas.append(valores)

    return pd.DataFrame(linhas[1:], columns=linhas[0])


df = ler_planilha_xml(arquivo)
quantidade_partes = 2

for i in range(1, quantidade_partes + 1):
    inicio = (i - 1) * len(df) // quantidade_partes
    fim = i * len(df) // quantidade_partes
    parte = df.iloc[inicio:fim]
    parte.to_excel(f"{arquivo.stem}_parte{i}.xlsx", index=False)
    print(f"Parte {i}: {len(parte)} linhas")

