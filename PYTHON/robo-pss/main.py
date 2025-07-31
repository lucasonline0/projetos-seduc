import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# Configura o Chrome
def iniciar_driver():
    options = Options()
    options.add_experimental_option("prefs", {
        "download.default_directory": os.path.abspath("downloads"),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    })
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

# Acessa e percorre os links até a página final
def navegar_e_baixar(driver):
    # Passo 1: Página inicial
    driver.get("https://www.seduc.pa.gov.br/sitenovo/seduc/")
    time.sleep(3)

    # Passo 2: Clicar no banner do PSS geral
    driver.find_element(By.XPATH, '//img[contains(@src, "5f333baa0c8883a29af33b5924d750cc")]').click()
    time.sleep(3)

    # Passo 3: Clicar no banner "PSS 04/2025 SECTET"
    driver.find_element(By.XPATH, '//img[contains(@src, "BANNER2%20(9)-afd65.png")]').click()
    time.sleep(3)

    # Passo 4: Baixar todos os arquivos disponíveis
    links = driver.find_elements(By.XPATH, '//a[contains(@href, "/site/public/upload/")]')

    print(f"Encontrados {len(links)} arquivos para download.")

    for link in links:
        url = link.get_attribute("href")
        driver.execute_script("window.open(arguments[0]);", url)
        time.sleep(2)

# Executa tudo
if __name__ == "__main__":
    os.makedirs("downloads", exist_ok=True)
    driver = iniciar_driver()
    try:
        navegar_e_baixar(driver)
    finally:
        print("Processo finalizado. Feche manualmente o navegador se necessário.")
