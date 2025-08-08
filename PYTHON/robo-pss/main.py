import os
import re
import time
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

DESKTOP = os.path.join(os.path.expanduser("~"), "Desktop")
BASE_DIR = os.path.join(DESKTOP, "PSS")

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

# ==== FUNÇÕES ====

def iniciar_driver():
    options = Options()
    options.add_argument("--headless=new")
    return webdriver.Chrome(options=options)

def formatar_nome_arquivo(nome):
    nome_sem_ext, ext = os.path.splitext(nome)
    nome_sem_ext = nome_sem_ext.replace("%", " ").replace("_", " ").replace("-", " ")
    nome_sem_ext = re.sub(r"\d+", "", nome_sem_ext)
    nome_sem_ext = re.sub(r"(?i)(\w)\1{1,}", r"\1", nome_sem_ext)
    nome_sem_ext = re.sub(r"\b\w\b", "", nome_sem_ext)
    nome_sem_ext = re.sub(r"\s{2,}", " ", nome_sem_ext)
    nome_final = nome_sem_ext.strip().title() + ext
    return nome_final

def formatar_nome_pasta(titulo_original):
    match = re.search(r"(PSS[\s_-]*\d+).*?-(.*?)\s*-", titulo_original, re.IGNORECASE)
    if match:
        pss = match.group(1).replace("_", " ").replace("-", " ").strip()
        cargo = match.group(2).strip().upper()
        cargo = re.sub(r"[^A-Z\s]", "", cargo)
        return f"{pss} - {cargo}"
    else:
        return re.sub(r"[\\/:*?\"<>|]", "", titulo_original).strip()

def baixar_arquivo(url, caminho_destino):
    try:
        r = requests.get(url)
        with open(caminho_destino, "wb") as f:
            f.write(r.content)
        print(f"✔️ Baixado: {os.path.basename(caminho_destino)}")
    except Exception as e:
        print(f"❌ Erro ao baixar {url}: {e}")

def processar_pagina(driver, url):
    try:
        driver.get(url)
        time.sleep(2)

        titulo = driver.title.strip()
        nome_pasta = formatar_nome_pasta(titulo)

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
            nome_original = href.split("/")[-1]
            nome_formatado = formatar_nome_arquivo(nome_original)

            destino = (
                os.path.join(pasta_especial, nome_formatado)
                if "CONVOCACAO" in nome_original.upper() and "ESPECIAL" in nome_original.upper()
                else os.path.join(pasta_destino, nome_formatado)
            )
            baixar_arquivo(href, destino)

        print(f"📁 Finalizado: {nome_pasta}\n")
    except Exception as e:
        print(f"❌ Erro ao processar {url}: {e}")

def baixar_todos_os_arquivos():
    os.makedirs(BASE_DIR, exist_ok=True)
    driver = iniciar_driver()

    for url in BANNERS_URLS:
        processar_pagina(driver, url)

    driver.quit()
    print(f"\n✅ Todos os arquivos foram baixados na sua Área de Trabalho em: {BASE_DIR}")

if __name__ == "__main__":
    baixar_todos_os_arquivos()
