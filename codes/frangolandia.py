Função principal para selecionar e ir com ChromeDrive.

import os
import re
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://frangolandia.com/encartes/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Concorrentes/Frangolandia"
os.makedirs(ENCARTE_DIR, exist_ok=True)

# ========= CHROME HEADLESS =========
def build_headless_chrome():
    options = webdriver.ChromeOptions()
# preferências (mantive as suas para PDF; não atrapalham, mesmo não sendo usadas aqui)
    prefs = {
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)

#     headless e flags de CI
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--lang=pt-BR,pt")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)

driver = build_headless_chrome()
wait = WebDriverWait(driver, 15)

def encontrar_data():
     #Exemplo de busca por textos de botões/labels na página (ajuste o seletor se quiser usar)
    try:
        enc_data = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//span[contains(@class, "elementor-button-text")]')))
    except Exception:
        return "sem_data"
    
    for div in enc_data:
        texto = div.text.strip()
        if texto:
            nome_pasta = re.sub(r'[\\/*?:"<>|\s]', '_', texto)
            return nome_pasta
    return "sem_data"

def initialazed():
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        })
    
#Procurar campanhas na Pagina HTML
def procurar_campanhas():
    print("Selecionando a campanha, data, mes e dia")
    try:
        time.sleep(2)
        data = driver.find_element(By.XPATH, "//span[contains(@class, 'elementor-button-text')]")
        data = data.text
        Select(data)
        campanha = driver.find_element(By.XPATH, "//h3[contains(@class, 'elementor-heading-title elementor-size-default')]")
        campanha.text
        Select(campanha)
    except:
        print("Campanhas e datas não encontradas")

#Select the elements and save in spreeadsheet       
def save_in_spreadsheet(campanha, data):
    try:
        print("Salvando na planilha")
        time.sleep(2)
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "Empresa", sheet["A2"] = "Frangolandia"
        sheet["B1"] = "Campanha", sheet["B2"] = campanha
        sheet["C1"] = "Data", sheet["C2"] = data
        workbook.save = ENCARTE_DIR/"campanhas_frangolandia.xlsx"
    except:
        print("Não foi possivel salvar")

try:
    initialazed()
    procurar_campanhas()
    save_in_spreadsheet()
    print("\n Processo finalizado.")
except Exception as e:
    print(f" Erro geral: {e}")
finally:
    driver.quit()
