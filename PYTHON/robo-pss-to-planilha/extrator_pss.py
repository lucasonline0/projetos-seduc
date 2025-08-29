import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os

def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
    return text

def extract_info_from_text(text, filename):
    data = []

    disciplinas_comuns = [
        "Matemática", "Língua Portuguesa", "História", "Geografia", "Educação Física",
        "Ciências", "Biologia", "Química", "Física", "Sociologia", "Filosofia",
        "Artes", "Inglês", "Espanhol", "Educação Artística", "Ensino Religioso",
        "Informática", "Agropecuária", "Zootecnia", "Enfermagem", "Contabilidade",
        "Administração", "Direito", "Pedagogia", "Letras", "Educação Especial",
        "Ciências Agrárias", "Tecnologias", "Educação do Campo", "EJA CAMPO",
        "Acompanhamento Especializado", "Intérprete de Libras", "Merendeira", "SOME",
        "Contador", "Professor"
    ]
    
    municipios_comuns = [
        "Belém", "Ananindeua", "Santarém", "Castanhal", "Marabá", "Parauapebas",
        "Altamira", "Redenção", "Paragominas", "Tucuruí", "Barcarena", "Abaetetuba",
        "Cametá", "Bragança", "Capanema", "Itaituba", "Marituba", "São Félix do Xingu",
        "Xinguara", "Tailândia", "Breves", "Portel", "Vigia", "Salinópolis",
        "Benevides", "Santa Izabel do Pará", "Capitão Poço", "Monte Alegre",
        "Oriximiná", "Dom Eliseu", "Ulianópolis", "Jacundá", "Tomé-Açu", "Conceição do Araguaia",
        "Cumaru do Norte", "Floresta do Araguaia", "Pau D\"Arco", "Santa Maria das Barreiras",
        "Santana do Araguaia", "Mãe do Rio", "Aurora do Pará", "Ipixuna do Pará",
        "Jacareacanga", "Novo Progresso", "Placas", "Rurópolis", "Trairão", "Curralinho",
        "Bagre", "Chaves", "Melgaço", "Gurupá", "Almeirim", "Medicilândia", "Porto de Moz",
        "Senador José Porfírio", "Uruará", "Vitória do Xingu", "Goianésia do Pará",
        "Novo Repartimento", "Pacajá", "Curua", "Faro", "Juruti", "Terra Santa",
        "Brasil Novo", "Anapu", "São Miguel do Guamá", "São João da Ponta", "São Francisco do Pará",
        "São Domingos do Capim", "Santa Maria do Pará", "Marapanim", "Inhangapi", "Curuçá",
        "São Caetano de Odivelas", "Santo Antônio do Tauá", "Santa Luzia do Pará",
        "Bonito", "Nova Timboteua", "Ourém", "Peixe-Boi", "Primavera", "Quatipuru",
        "São João de Pirabas", "Augusto Corrêa", "Cachoeira do Arari", "Colares", "Mocajuba",
        "Mojú", "Nova Esperança do Piriá", "Oeiras do Pará", "Ponta de Pedras", "São Sebastião da Boa Vista",
        "Soure", "Viseu", "Acará", "Afuá", "Água Azul do Norte", "Alenquer", "Anajás", "Aveiro",
        "Baião", "Bannach", "Belterra", "Bom Jesus do Tocantins", "Bonito", "Buarque", "Cachoeira do Piriá",
        "Canaã dos Carajás", "Canaã dos Carajás", "Capanema", "Castelo dos Sonhos", "Chaves",
        "Colares", "Conceição do Araguaia", "Concórdia do Pará", "Curuá", "Curionópolis",
        "Curralinho", "Garrafão do Norte", "Goianésia do Pará", "Gurupá", "Igarapé-Açu",
        "Igarapé-Miri", "Inhangapi", "Ipixuna do Pará", "Irituia", "Itupiranga", "Jacareacanga",
        "Jacundá", "Juruti", "Limoeiro do Ajuru", "Magalhães Barata", "Maracanã", "Marapanim",
        "Medicilândia", "Melgaço", "Mocajuba", "Moju", "Monte Alegre", "Muaná", "Nova Esperança do Piriá",
        "Nova Ipixuna", "Nova Timboteua", "Novo Progresso", "Novo Repartimento", "Óbidos",
        "Oeiras do Pará", "Ourem", "Ourilândia do Norte", "Pacajá", "Palestina do Pará",
        "Paragominas", "Pau D\"Arco", "Peixe-Boi", "Ponta de Pedras", "Portel", "Porto de Moz",
        "Primavera", "Quatipuru", "Redenção", "Rio Maria", "Rondon do Pará", "Rurópolis",
        "Salinópolis", "Santa Bárbara do Pará", "Santa Cruz do Arari", "Santa Izabel do Pará",
        "Santa Luzia do Pará", "Santa Maria das Barreiras", "Santa Maria do Pará", "Santana do Araguaia",
        "Santarém Novo", "Santo Antônio do Tauá", "São Caetano de Odivelas", "São Domingos do Araguaia",
        "São Domingos do Capim", "São Francisco do Pará", "São Geraldo do Araguaia",
        "São João da Ponta", "São João de Pirabas", "São João do Araguaia", "São Miguel do Guamá",
        "São Sebastião da Boa Vista", "Sapucaia", "Soure", "Tailândia", "Terra Alta", "Terra Santa",
        "Tomé-Açu", "Tracuateua", "Trairão", "Tucumã", "Tucuruí", "Ulianópolis", "Uruará",
        "Vigia", "Viseu", "Vitória do Xingu", "Xinguara"
    ]

    pss = "N/A"
    year = "N/A"

    pss_file_match = re.search(r'PSS(?:\s*Nº?)?(\s*\d{2}[-/]\d{4}|\s*\d{4})', filename, re.IGNORECASE)
    if pss_file_match:
        pss = pss_file_match.group(1).strip().replace('-', '/')
        if len(pss) == 4:
            pss = f"00/{pss}"

    # Tentar extrair PSS do texto
    if pss == "N/A":
        pss_text_match = re.search(r'PROCESSO SELETIVO SIMPLIFICADO Nº?\s*(\d{2}[-/]\d{4}|\d{4})', text, re.IGNORECASE)
        if pss_text_match:
            pss = pss_text_match.group(1).strip().replace('-', '/')
            if len(pss) == 4: 
                pss = f"00/{pss}"
        else:
            pss_text_match = re.search(r'PSS\s*Nº?\s*(\d{2}[-/]\d{4}|\d{4})', text, re.IGNORECASE)
            if pss_text_match:
                pss = pss_text_match.group(1).strip().replace('-', '/')
                if len(pss) == 4: 
                    pss = f"00/{pss}"

    year_file_match = re.search(r'(\d{4})', filename)
    if year_file_match:
        year = year_file_match.group(1)

    if year == "N/A":
        year_text_match = re.search(r'(\d{4})', text)
        if year_text_match:
            year = year_text_match.group(1)

    quadro_vagas_match = re.search(r'QUADRO DE VAGAS\s*(.*?)(?=(?:RELAÇÃO DE E-MAIL|TOTAL DE VAGAS|\Z))', text, re.DOTALL | re.IGNORECASE)
    
    if quadro_vagas_match:
        quadro_vagas_text = quadro_vagas_match.group(1)
        lines = quadro_vagas_text.split('\n')
        
        found_municipalities = []
        found_disciplines = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            for mun in municipios_comuns:
                if re.search(r'\b' + re.escape(mun) + r'\b', line, re.IGNORECASE):
                    found_municipalities.append(mun)

            for disc in disciplinas_comuns:
                if re.search(r'\b' + re.escape(disc) + r'\b', line, re.IGNORECASE):
                    found_disciplines.append(disc)

        found_municipalities = list(set(found_municipalities)) if found_municipalities else [None]
        found_disciplines = list(set(found_disciplines)) if found_disciplines else [None]

        for m in found_municipalities:
            for d in found_disciplines:
                data.append({
                    "Ano": year,
                    "PSS": pss,
                    "Edital": f"Edital nº {filename}",
                    "Município": m,
                    "Disciplina": d,
                    "Arquivo_Origem": filename,
                    "Data_Extração": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
    else:
        found_municipalities = []
        found_disciplines = []

        for mun in municipios_comuns:
            if re.search(r'\b' + re.escape(mun) + r'\b', text, re.IGNORECASE):
                found_municipalities.append(mun)
        
        for disc in disciplinas_comuns:
            if re.search(r'\b' + re.escape(disc) + r'\b', text, re.IGNORECASE):
                found_disciplines.append(disc)

        found_municipalities = list(set(found_municipalities)) if found_municipalities else [None]
        found_disciplines = list(set(found_disciplines)) if found_disciplines else [None]

        for m in found_municipalities:
            for d in found_disciplines:
                data.append({
                    "Ano": year,
                    "PSS": pss,
                    "Edital": f"Edital nº {filename}",
                    "Município": m,
                    "Disciplina": d,
                    "Arquivo_Origem": filename,
                    "Data_Extração": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })

    return data

PASTA_PDFS = "C:/Users/SEDUC/Desktop/conv especiais"
ARQUIVO_SAIDA = "convocacoes_especiais.xlsx"

all_data = []

for arquivo in os.listdir(PASTA_PDFS):
    if not arquivo.lower().endswith(".pdf"):
        continue

    caminho = os.path.join(PASTA_PDFS, arquivo)
    print(f"📄 Processando PDF: {arquivo}")

    texto_pdf = extract_text_from_pdf(caminho)
    extracted_data = extract_info_from_text(texto_pdf, arquivo)
    
    if extracted_data:
        all_data.extend(extracted_data)
        print(f"✅ {len(extracted_data)} dados extraídos deste PDF.\n")
    else:
        print(f"⚠️ Nenhum dado extraído deste PDF: {arquivo}.\n")

if all_data:
    df = pd.DataFrame(all_data)
    df.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"🎉 Extração concluída! Dados salvos em {ARQUIVO_SAIDA}")
else:
    print("❌ Nenhuma informação extraída de nenhum PDF.")