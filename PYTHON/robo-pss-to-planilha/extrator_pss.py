import pdfplumber
import pandas as pd
import re
from datetime import datetime
import os

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    lines = page_text.split("\n")
                    possible_footer_keywords = [r'\bTel\b', r'\bFax\b', r'\bCEP\b', r'Rodovia', r'Km\s*\d+']
                    
                    num_lines_to_check = 4
                    lines_to_remove = 0
                    for i in range(1, num_lines_to_check + 1):
                        if len(lines) >= i:
                            line_to_check = lines[-i]
                            if any(re.search(keyword, line_to_check, re.IGNORECASE) for keyword in possible_footer_keywords):
                                lines_to_remove = i
                    
                    if lines_to_remove > 0:
                        lines = lines[:-lines_to_remove]

                    text += "\n".join(lines) + "\n"
    except Exception as e:
        print(f"Erro ao ler o PDF {os.path.basename(pdf_path)}: {e}")
    return text

def extract_info_from_text(text, filename):
    data = []

    disciplinas_comuns = sorted(list(set([
        "Matemática", "Língua Portuguesa", "História", "Geografia", "Educação Física",
        "Ciências", "Biologia", "Química", "Física", "Sociologia", "Filosofia",
        "Artes", "Inglês", "Espanhol", "Educação Artística", "Ensino Religioso",
        "Informática", "Agropecuária", "Zootecnia", "Enfermagem", "Contabilidade",
        "Administração", "Direito", "Pedagogia", "Letras", "Educação Especial",
        "Ciências Agrárias", "Tecnologias", "Educação do Campo", "EJA CAMPO",
        "Acompanhante Especializado", "Intérprete de Libras", "Merendeira", "SOME",
        "Contador", "Professor", "Qualquer disciplina", "Educação Geral", "Língua Inglesa",
        "Apoio - Tradutor Intérprete de Libras", "Atendimento Educacional Especializado", "AEE",
        "Brailista", "Ciências da Natureza e suas Tecnologias", "Ciências Humanas e suas Tecnologias",
        "Linguagens Códigos e suas Tecnologias", "Professor Coordenador(a) de Turma", "Projeto SEI"
    ])))
    
    municipios_comuns = sorted(list(set([
        "Abel Figueiredo", "Acará", "Afuá", "Água Azul do Norte", "Alenquer", "Almeirim", "Altamira",
        "Anajás", "Ananindeua", "Anapu", "Augusto Corrêa", "Aurora do Pará", "Aveiro", "Bagre",
        "Baião", "Bannach", "Barcarena", "Belém", "Belterra", "Benevides", "Bom Jesus do Tocantins",
        "Bonito", "Bragança", "Brasil Novo", "Brejo Grande do Araguaia", "Breu Branco", "Buarque",
        "Cachoeira do Arari", "Cachoeira do Piriá", "Cametá", "Canaã dos Carajás", "Capanema",
        "Capitão Poço", "Castanhal", "Chaves", "Colares", "Conceição do Araguaia", "Concórdia do Pará",
        "Cumaru do Norte", "Curuá", "Curuçá", "Curionópolis", "Curralinho", "Dom Eliseu",
        "Eldorado dos Carajás", "Faro", "Floresta do Araguaia", "Garrafão do Norte", "Goianésia do Pará",
        "Gurupá", "Icoaraci", "Igarapé-Açu", "Igarapé-Miri", "Inhangapi", "Ipixuna do Pará", "Irituia",
        "Itaituba", "Itupiranga", "Jacareacanga", "Jacundá", "Juruti", "Limoeiro do Ajuru",
        "Mãe do Rio", "Magalhães Barata", "Marabá", "Maracanã", "Marapanim", "Marituba",
        "Medicilândia", "Melgaço", "Mocajuba", "Moju", "Mojuí dos Campos", "Monte Alegre", "Muaná",
        "Nova Esperança do Piriá", "Nova Ipixuna", "Nova Timboteua", "Novo Progresso", "Novo Repartimento",
        "Óbidos", "Oeiras do Pará", "Ourém", "Ourilândia do Norte", "Pacajá", "Palestina do Pará",
        "Paragominas", "Parauapebas", "Pau D\'Arco", "Peixe-Boi", "Piçarra", "Placas", "Ponta de Pedras",
        "Portel", "Porto de Moz", "Prainha", "Primavera", "Quatipuru", "Redenção", "Rio Maria",
        "Rondon do Pará", "Rurópolis", "Salinópolis", "Salvaterra", "Santa Bárbara do Pará",
        "Santa Cruz do Arari", "Santa Izabel do Pará", "Santa Luzia do Pará", "Santa Maria das Barreiras",
        "Santa Maria do Pará", "Santana do Araguaia", "Santarém", "Santarém Novo", "Santo Antônio do Tauá",
        "São Caetano de Odivelas", "São Domingos do Araguaia", "São Domingos do Capim",
        "São Félix do Xingu", "São Francisco do Pará", "São Geraldo do Araguaia", "São João da Ponta",
        "São João de Pirabas", "São João do Araguaia", "São Miguel do Guamá", "São Sebastião da Boa Vista",
        "Sapucaia", "Senador José Porfírio", "Soure", "Tailândia", "Terra Alta", "Terra Santa",
        "Tomé-Açu", "Tracuateua", "Trairão", "Tucumã", "Tucuruí", "Ulianópolis", "Uruará", "Vigia",
        "Viseu", "Vitória do Xingu", "Xinguara", "Anexo Estrela do Pará", "Bela Vista do Baixo",
        "Bela Vista do Mocoões", "Cachoeira da Serra", "Casa de Tábua", "Castelo dos Sonhos",
        "Comunidade Arrozal", "Comunidade do Ipanema", "Comunidade Maria Ribeira",
        "Comunidade Nossa Senhora de Nazaré", "Comunidade Porto Alegre", "Comunidade Santa Luzia",
        "Comunidade São Francisco", "Comunidade São José", "Comunidade São Pedro",
        "Distrito de Moraes Almeida", "Distrito de Mosqueiro", "Estrela do Pará", "Forquilha",
        "Mata Verde", "Santa Maria de Icatu", "Tabatinga", "Vila Bela Vista", "Vila do Carmo",
        "Vila Estrela do Pará", "Vila Lawton", "Vila Martins Ferreira", "Vila Nazaré", "Vila Nova",
        "Vila Ponta de Pedras", "Vila Romaria", "Vila Santa Cruz"
    ])))

    pss = "N/A"
    year = "N/A"

    pss_match = re.search(r'PSS\s*(?:Nº)?\s*(\d{2,4}/\d{4})', text, re.IGNORECASE) or \
                re.search(r'PROCESSO SELETIVO SIMPLIFICADO\s*(?:Nº)?\s*(\d{2,4}/\d{4})', text, re.IGNORECASE)
    if pss_match:
        pss = pss_match.group(1)

    date_match = re.search(r'Belém,\s*\d{1,2}\s+de\s+\w+\s+de\s+(\d{4})', text, re.IGNORECASE)
    if date_match:
        year = date_match.group(1)
    else:
        year_match_in_edital = re.search(r'EDITAL\s*(?:nº)?\s*\d{1,3}/(\d{4})', text, re.IGNORECASE)
        if year_match_in_edital:
            year = year_match_in_edital.group(1)
        else:
            valid_years = "|".join(map(str, range(2019, 2026)))
            fallback_match = re.search(r'\b(' + valid_years + r')\b', text + " " + filename)
            if fallback_match:
                year = fallback_match.group(1)

    def find_entries(text_to_search):
        municipalities = {mun for mun in municipios_comuns if re.search(r'\b' + re.escape(mun) + r'\b', text_to_search, re.IGNORECASE)}
        disciplines = {disc for disc in disciplinas_comuns if re.search(r'\b' + re.escape(disc) + r'\b', text_to_search, re.IGNORECASE)}
        return municipalities, disciplines

    quadro_vagas_match = re.search(r'QUADRO DE VAGAS(.*?)(?:TOTAL DE VAGAS|RELAÇÃO DE E-MAIL|Tiago Lima e Silva)', text, re.DOTALL | re.IGNORECASE)
    
    found_municipalities, found_disciplines = set(), set()
    if quadro_vagas_match:
        quadro_text = quadro_vagas_match.group(1)
        found_municipalities, found_disciplines = find_entries(quadro_text)

    if not found_municipalities and not found_disciplines:
        found_municipalities, found_disciplines = find_entries(text)

    final_municipalities = list(found_municipalities) if found_municipalities else ["N/A"]
    final_disciplines = list(found_disciplines) if found_disciplines else ["N/A"]

    for m in final_municipalities:
        for d in final_disciplines:
            data.append({
                "Ano": year,
                "PSS": pss,
                "Edital": filename.replace('.pdf', ''),
                "Município": m,
                "Disciplina": d,
                "Arquivo_Origem": filename,
                "Data_Extração": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

    return data

PASTA_PDFS = "C:/Users/SEDUC/Desktop/conv especiais"
ARQUIVO_SAIDA = "convocacoes_especiais.xlsx"

all_data = []

if not os.path.isdir(PASTA_PDFS):
    print(f" Erro: A pasta \033[91m'{PASTA_PDFS}'\033[0m não foi encontrada.")
else:
    pdf_files = [f for f in os.listdir(PASTA_PDFS) if f.lower().endswith(".pdf")]
    
    if not pdf_files:
        print(f"Nenhum arquivo PDF encontrado em \033[93m'{PASTA_PDFS}'\033[0m.")
    else:
        for arquivo in pdf_files:
            caminho = os.path.join(PASTA_PDFS, arquivo)
            print(f"\033[94m📄 Processando:\033[0m {arquivo}")

            texto_pdf = extract_text_from_pdf(caminho)
            
            if not texto_pdf:
                print(f"   -> \033[91mNão foi possível extrair texto deste PDF.\033[0m\n")
                continue

            extracted_data = extract_info_from_text(texto_pdf, arquivo)
            
            if extracted_data:
                all_data.extend(extracted_data)
                print(f"   -> \033[92m✅ {len(extracted_data)} registros extraídos.\033[0m\n")
            else:
                print(f"   -> \033[93mNenhum dado relevante encontrado.\033[0m\n")

if all_data:
    df = pd.DataFrame(all_data)
    df.drop_duplicates(subset=["Ano", "PSS", "Edital", "Município", "Disciplina"], inplace=True)
    df.to_excel(ARQUIVO_SAIDA, index=False)
    print(f"\033[92mExtração concluída! {len(df)} registros únicos salvos em {ARQUIVO_SAIDA}\033[0m")
else:
    print("\033[91m❌ Nenhuma informação foi extraída de nenhum dos PDFs.\033[0m")
