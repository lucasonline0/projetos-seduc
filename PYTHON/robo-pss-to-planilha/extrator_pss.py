import os
import re
import datetime
import pandas as pd
import pdfplumber
import unicodedata
from itertools import product

PDF_FOLDER_PATH = r"C:/Users/SEDUC/Desktop/conv especiais"
OUTPUT_FILE_NAME = "convocacoes_especiais.xlsx"

# Lista de municípios e distritos a serem procurados nos textos dos editais.
MUNICÍPIOS = [
    'Abaetetuba', 'Abel Figueiredo', 'Acará', 'Afuá', 'Agua Azul do Norte',
    'Alenquer', 'Almeirim', 'Altamira', 'Anajás', 'Ananindeua', 'Anapu',
    'Augusto Corrêa', 'Aurora do Pará', 'Aveiro', 'Bagre', 'Baião', 'Bannach',
    'Barcarena', 'Belém', 'Belterra', 'Benevides', 'Bom Jesus do Tocantins',
    'Bonito', 'Bragança', 'Brasil Novo', 'Brejo Grande do Araguaia',
    'Breu Branco', 'Breves', 'Bujaru', 'Cachoeira do Arari', 'Cachoeira do Piriá',
    'Cametá', 'Canaa dos Carajas', 'Capanema', 'Capitão Poço', 'Castanhal',
    'Chaves', 'Colares', 'Conceição do Araguaia', 'Concórdia do Pará',
    'Cumaru do Norte', 'Curionopolis', 'Curralinho', 'Curuá', 'Curuçá',
    'Dom Eliseu', 'Eldorado dos Carajas', 'Faro', 'Floresta do Araguaia',
    'Garrafão do Norte', 'Goianésia do Pará', 'Gurupá', 'Icoaraci', 'Igarapé-Açu',
    'Igarapé Miri', 'Inhangapi', 'Ipixuna do Pará', 'Irituia', 'Itaituba',
    'Itupiranga', 'Jacareacanga', 'Jacundá', 'Juruti', 'Limoeiro do Ajuru',
    'Mãe do Rio', 'Magalhães Barata', 'Marabá', 'Maracanã', 'Marapanim',
    'Marituba', 'Medicilândia', 'Melgaço', 'Mocajuba', 'Moju', 'Mojuí dos Campos',
    'Monte Alegre', 'Monte Dourado', 'Mosqueiro', 'Muaná',
    'Nova Esperança do Piriá', 'Nova Ipixuna', 'Nova Timboteua',
    'Novo Progresso', 'Novo Repartimento', 'Obidos', 'Oeiras do Pará',
    'Oriximiná', 'Ourém', 'Ourilandia do Norte', 'Pacajá', 'Palestina do Pará',
    'Paragominas', 'Parauapebas', 'Pau D\'Arco', 'Peixe-boi', 'Piçarra', 'Placas',
    'Ponta de Pedras', 'Portel', 'Porto de Moz', 'Prainha', 'Primavera',
    'Quatipuru', 'Redenção', 'Rio Maria', 'Rondon do Pará', 'Rurópolis',
    'Salinas', 'Salinópolis', 'Salvaterra', 'Santa Barbara', 'Santa Cruz da Arari',
    'Santa Izabel do Pará', 'Santa Luzia do Pará', 'Santa Maria das Barreiras',
    'Santa Maria do Pará', 'Santana do Araguaia', 'Santarém', 'Santarém Novo',
    'Santo Antônio do Pará', 'Santo Antônio do Tauá', 'São Caetano de Odivelas',
    'São Domingos do Araguaia', 'São Domingos do Capim', 'São Felix do Xingu',
    'São Francisco do Pará', 'São Geraldo do Araguaia', 'São João da Ponta',
    'São João de Pirabas', 'São João do Araguaia', 'São Miguel do Guamá',
    'São Sebastiao da Boa Vista', 'Sapucaia', 'Senador José Porfírio', 'Soure',
    'Tailandia', 'Terra Alta', 'Terra Santa', 'Tomé-Açu', 'Tracuateua', 'Trairão',
    'Tucumã', 'Tucuruí', 'Ulianópolis', 'Uruará', 'Vigia', 'Viseu',
    'Vitória do Xingu', 'Xinguara'
]

DISCIPLINAS = [
    'Acompanhante Especializado', 'APOIO-TRADUTOR INTERPRETE DE LIBRAS', 'Artes',
    'Biologia', 'BRAILISTA', 'CIÊNCIAS AGRÁRIAS E SUAS TECNOLOGIAS',
    'CIÊNCIAS HUMANAS E SUAS TECNOLOGIAS', 'CIÊNCIAS NATURAIS E SUAS TECNOLOGIAS',
    'CONTADOR', 'DESIGN', 'Educação Física', 'Educação Geral', 'Espanhol',
    'Filosofia', 'Fisica', 'Geografia', 'Historia', 'Ingles',
    'Intérprete de Libras', 'LINGUAGENS CÓDIGOS E SUAS TECNOLOGIAS',
    'Lingua Inglesa', 'Lingua Portuguesa',
    'MATEMÁTICA, CÓDIGOS E SUAS TECNOLOGIAS', 'Matemática', 'Português',
    'PROFESSOR COORDENADOR(A) DE TURMA',
    'PROFESSOR DE ATENDIMENTO EDUCACIONAL ESPECIALIZADO AEE',
    'Professor Educador Agrárias',
    'Professor Educador Ciências Humanas e suas Tecnologias',
    'Professor Educador Ciências Naturais e suas Tecnologias',
    'Professor Educador Matemáticas e suas Tecnologias', 'Quimica', 'Sociologia'
]

def normalizar_texto(texto):
    if not texto:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return texto_sem_acento.upper().strip()

def extrair_texto_de_pdf(caminho_pdf):
    """
    Extrai todo o texto de um arquivo PDF.
    """
    texto_completo = ""
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for pagina in pdf.pages:
                texto = pagina.extract_text()
                if texto:
                    texto_completo += texto + "\n"
    except Exception as e:
        print(f"  [ERRO] Não foi possível ler o arquivo {os.path.basename(caminho_pdf)}: {e}")
    return texto_completo

def extrair_ano(texto, nome_arquivo):
    date_match = re.search(r'Belém,\s*\d{1,2}\s+de\s+\w+\s+de\s+(\d{4})', texto, re.IGNORECASE)
    if date_match:
        return date_match.group(1)

    year_match_in_edital = re.search(r'EDITAL\s*(?:nº)?\s*\d{1,3}/(\d{4})', texto, re.IGNORECASE)
    if year_match_in_edital:
        return year_match_in_edital.group(1)

    valid_years = "|".join(map(str, range(2019, datetime.datetime.now().year + 2)))
    fallback_match = re.search(r'\b(' + valid_years + r')\b', texto + " " + nome_arquivo)
    if fallback_match:
        return fallback_match.group(1)
        
    return "N/A"

def extrair_pss(texto):
    match = re.search(r'PROCESSO SELETIVO SIMPLIFICADO\s*(?:Nº)?\s*(\d{2,4}/\d{4})', texto, re.IGNORECASE) or \
            re.search(r'PSS\s*(?:Nº)?\s*(\d{2,4}/\d{4})', texto, re.IGNORECASE)
    
    return match.group(1) if match else "N/A"

def extrair_edital(texto):
    padrao = re.search(r"EDITAL\s*n?º?\s*([\d/]+)", texto, re.IGNORECASE)
    return padrao.group(1).strip() if padrao else "N/A"

def encontrar_entidades(texto, lista_entidades):
    encontrados = set()
    texto_normalizado = normalizar_texto(texto)
    for entidade in lista_entidades:
        entidade_normalizada = normalizar_texto(entidade)
        if re.search(r'\b' + re.escape(entidade_normalizada) + r'\b', texto_normalizado):
            encontrados.add(entidade)
    return sorted(list(encontrados))

def main():
    print("Iniciando a extração de dados dos editais da SEDUC...")
    
    if "COLOQUE_O_CAMINHO_DA_SUA_PASTA_AQUI" in PDF_FOLDER_PATH:
        print("[ERRO] Por favor, configure a variável 'PDF_FOLDER_PATH' no script com o caminho da sua pasta de PDFs.")
        return

    if not os.path.isdir(PDF_FOLDER_PATH):
        print(f"[ERRO] O diretório especificado não foi encontrado: {PDF_FOLDER_PATH}")
        return

    arquivos_pdf = [f for f in os.listdir(PDF_FOLDER_PATH) if f.lower().endswith(".pdf")]

    if not arquivos_pdf:
        print(f"[AVISO] Nenhum arquivo PDF encontrado no diretório: {PDF_FOLDER_PATH}")
        return

    dados_extraidos = []
    data_extracao_atual = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    print(f"Encontrados {len(arquivos_pdf)} arquivos PDF para processar.")

    for nome_arquivo in arquivos_pdf:
        caminho_completo = os.path.join(PDF_FOLDER_PATH, nome_arquivo)
        print(f"\nProcessando arquivo: {nome_arquivo}...")

        texto_pdf = extrair_texto_de_pdf(caminho_completo)
        if not texto_pdf:
            continue

        # Extração dos dados principais
        ano = extrair_ano(texto_pdf, nome_arquivo)
        pss = extrair_pss(texto_pdf)
        edital = extrair_edital(texto_pdf)
        
        # Busca por municípios sempre
        municipios_encontrados = encontrar_entidades(texto_pdf, MUNICÍPIOS)
        
        # VERIFICA O NOME DO ARQUIVO PARA "PROJETO SEI"
        if "projeto sei" in nome_arquivo.lower():
            # Se encontrar, define a disciplina manualmente e pula a busca padrão
            disciplinas_encontradas = ["PROJETO SEI"]
            print("  -> Regra especial ativada: Arquivo identificado como 'PROJETO SEI'.")
        else:
            # Se não, continua com a busca normal por disciplinas
            disciplinas_encontradas = encontrar_entidades(texto_pdf, DISCIPLINAS)
        # fim da logica

        print(f"  - PSS: {pss}, Edital: {edital}, Ano: {ano}")
        print(f"  - Municípios encontrados: {len(municipios_encontrados)} -> {municipios_encontrados if municipios_encontrados else 'Nenhum'}")
        print(f"  - Disciplinas encontradas: {len(disciplinas_encontradas)} -> {disciplinas_encontradas if disciplinas_encontradas else 'Nenhuma'}")

        if not municipios_encontrados and not disciplinas_encontradas:
             municipios_para_iterar = ["N/A"]
             disciplinas_para_iterar = ["N/A"]
        else:
             municipios_para_iterar = municipios_encontrados if municipios_encontrados else ["N/A"]
             disciplinas_para_iterar = disciplinas_encontradas if disciplinas_encontradas else ["N/A"]

        for municipio, disciplina in product(municipios_para_iterar, disciplinas_para_iterar):
            linha_dados = {
                "Ano": ano,
                "PSS": pss,
                "Edital": edital,
                "Município": municipio,
                "Disciplina_Função": disciplina,
                "Arquivo_Origem": nome_arquivo,
                "Data_Extração": data_extracao_atual
            }
            dados_extraidos.append(linha_dados)

    if not dados_extraidos:
        print("Nenhum dado foi extraído. Verifique o conteúdo dos arquivos PDF.")
        return

    # Criação do DataFrame e salvamento em Excel
    df = pd.DataFrame(dados_extraidos)
    
    try:
        df.to_excel(OUTPUT_FILE_NAME, index=False, engine='openpyxl')
        print(f"\n\n[SUCESSO] Processo concluído!")
        print(f"Os dados foram salvos no arquivo: {os.path.abspath(OUTPUT_FILE_NAME)}")
    except Exception as e:
        print(f"\n[ERRO] Não foi possível salvar o arquivo Excel: {e}")

if __name__ == "__main__":
    main()