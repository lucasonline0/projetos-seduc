import os
import time
import re
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Detecta Área de Trabalho do sistema
DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
BASE_DIR = os.path.join(DESKTOP, "downloads_pss")

# Lista de links dos banners (processos seletivos)
BANNERS_URLS = [
    "https://www.seduc.pa.gov.br/pagina/13936-pss-04-2025-sectet---professor-nivel-superior-e-especialista-em-educacao",
    "https://www.seduc.pa.gov.br/pagina/13820-pss-03-2025---analista-e-assistente-administrativo",
    "https://www.seduc.pa.gov.br/pagina/13686-pssq-02-2025---professor-do-ensino-regular---educacao-quilombola",
    "https://www.seduc.pa.gov.br/pagina/13543-pss-03-2024---professor-de-educacao-basica",
    "https://www.seduc.pa.gov.br/pagina/12983-pss-02-2024---analista-de-gestao-governamental-e-politica-educacional",
    "https://www.seduc.pa.gov.br/pagina/12870-pss-01-2024---professor-do-ensino-regular---libras-e-atendimento-especializado",
    "https://www.seduc.pa.gov.br/pagina/12848-pss-01-2023---sectet",
    "https://www.seduc.pa.gov.br/pagina/12688-pss-05-2023---educacao-especial",
    "https://www.seduc.pa.gov.br/pagina/12277-pss-04-2023---engenheiros--arquitetos-e-contadores",
    "https://www.seduc.pa.gov.br/pagina/12187-pss-03-2023---professores-ou-coordenadores-de-turma---eja-campo",
    "https://www.seduc.pa.gov.br/pagina/12161-pss-02-2023---psicologos-e-assistentes-sociais",
    "https://www.seduc.pa.gov.br/pagina/12028-pss-01-2023---professores-e-merendeiras",
    "https://www.seduc.pa.gov.br/pagina/11803-pss-03-2022-engenheiros-e-arquitetos",
    "https://www.seduc.pa.gov.br/pagina/11772-pss-002-2022---tradutor-interprete-de-libras",
    "https://www.seduc.pa.gov.br/pagina/11760-pss-001-2022---sectec---seduc---professor-e-tec-nivel-superior",
    "https://www.seduc.pa.gov.br/pagina/11455-pss-03-2021---tecnico-em-gestao-publica",
    "https://www.seduc.pa.gov.br/pagina/10646-pss-04-2020---professor-ensino-regular"
]

# Função para baixar um arquivo PDF individualmente
def baixar_arquivo(url, caminho_destino):
    try:
        r = requests.get(url)
        with open(caminho_destino, "wb") as f:
            f.write(r.content)
        print(f"✔️ Baixado: {os.path.basename(caminho_destino)}")
    except Exception as e:
        print(f"❌ Erro ao baixar {url}: {e}")

# Função para baixar todos os arquivos de uma página específica
def processar_pagina(driver, url):
    try:
        driver.get(url)
        time.sleep(2)

        titulo = driver.title.strip()
        nome_pasta = re.sub(r"[^\w\s-]", "", titulo).strip().replace(" ", "_")
        pasta_destino = os.path.join(BASE_DIR, nome_pasta)
        pasta_especial = os.path.join(pasta_destino, "CONVOCACOES_ESPECIAIS")
        os.makedirs(pasta_destino, exist_ok=True)
        os.makedirs(pasta_especial, exist_ok=True)

        links_pdf = driver.find_elements(By.XPATH, "//a[contains(@href, '.pdf')]")
        if not links_pdf:
            print(f"⚠️ Nenhum PDF encontrado em: {url}")
            return

        for link in links_pdf:
            href = link.get_attribute("href")
            nome_arquivo = href.split("/")[-1]

            destino = (
                os.path.join(pasta_especial, nome_arquivo)
                if "CONVOCACAO" in nome_arquivo.upper() and "ESPECIAL" in nome_arquivo.upper()
                else os.path.join(pasta_destino, nome_arquivo)
            )
            baixar_arquivo(href, destino)

        print(f"📁 Finalizado: {nome_pasta}\n")
    except Exception as e:
        print(f"❌ Erro ao processar {url}: {e}")

# Função principal para baixar todos os arquivos de todos os banners
def baixar_todos_os_arquivos():
    os.makedirs(BASE_DIR, exist_ok=True)

    options = Options()
    options.add_argument("--headless=new")
    driver = webdriver.Chrome(options=options)

    for url in BANNERS_URLS:
        processar_pagina(driver, url)

    driver.quit()
    print(f"\n✅ Todos os arquivos foram baixados na sua Área de Trabalho em: {BASE_DIR}")

# --- Execução direta (se rodar o script diretamente) ---
if __name__ == "__main__":
    baixar_todos_os_arquivos()
