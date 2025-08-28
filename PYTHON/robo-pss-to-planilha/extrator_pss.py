import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os

PASTA_PDFS = r"C:\Users\SEDUC\Desktop\conv especiais"
ARQUIVO_SAIDA = "convocacoes_especiais.xlsx"

municipios_para = ["Belém","Abaetetuba","Santarém","Castanhal","Marabá"]  # lista exemplo
disciplinas_possiveis = ["Matemática","Língua Portuguesa","História","Geografia","Educação Física"]

dados = []

for arquivo in os.listdir(PASTA_PDFS):
    if not arquivo.lower().endswith(".pdf"):
        continue

    caminho = os.path.join(PASTA_PDFS, arquivo)
    print(f"📄 Processando PDF: {arquivo}")

    with pdfplumber.open(caminho) as pdf:
        texto = "\n".join([pagina.extract_text() for pagina in pdf.pages if pagina.extract_text()])

    encontrados_municipios = [m for m in municipios_para if m.lower() in texto.lower()]
    if not encontrados_municipios:
        encontrados_municipios = [None]

    encontrados_disciplinas = [d for d in disciplinas_possiveis if d.lower() in texto.lower()]
    if not encontrados_disciplinas:
        encontrados_disciplinas = [None]

    registros_pdf = 0
    for m in encontrados_municipios:
        for d in encontrados_disciplinas:
            dados.append({
                "Ano": "2024",
                "PSS": "PSS 05/2024",
                "Edital": f"Edital nº 05/2024 - {arquivo}",
                "Município": m,
                "Disciplina": d,
                "Arquivo_Origem": arquivo,
                "Data_Extração": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            registros_pdf += 1

    print(f"✅ {registros_pdf} dados extraídos deste PDF.\n")

df = pd.DataFrame(dados)
df.to_excel(ARQUIVO_SAIDA, index=False)
print(f"🎉 Extração concluída! Dados salvos em {ARQUIVO_SAIDA}")
