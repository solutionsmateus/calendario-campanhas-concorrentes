#Função principal para selecionar e ir com ChromeDrive.

import os
import re
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

#Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://novoatacarejo.com/oferta/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Concorrentes"
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

# === CHROME HEADLESS ===
def build_headless_chrome(download_dir: Path):
    options = webdriver.ChromeOptions()
    # headless e flags de CI
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
    # preferências de download (mantidas)
    prefs = {
        "download.prompt_for_download": False,
        "download.default_directory": str(download_dir),
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

driver = build_headless_chrome(ENCARTE_DIR)
wait = WebDriverWait(driver, 20)

def encontrar_data():
    # "h6 - TEXT LOCATION OF DATES IN PAGE"
    try:
        enc_data = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//h6[contains(\"TEXT\")]"))
        )
    except Exception:
        return "sem_data"
    for div in enc_data:
        texto = div.text.strip()
        if texto:
            nome_pasta = re.sub(r'[\\/*?:"<>|\s]', '_', texto)
            return nome_pasta
    return "sem_data"

def selecionar_loja():
    driver.get(BASE_URL)
    select_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select.select")))
    Select(select_element).select_by_visible_text("Olinda")
    time.sleep(4)  # aguarda carregar os tabloids

selecionar_loja()

def procurar_campanhas():
    campanha_and_data = driver.find_element(By.XPATH, "h6//contains('text')")
    campanha_and_data.text
    Select(campanha_and_data)

def save_in_spreadsheet(campanha_and_data):
    try:
        print("Salvando na planilha")
        time.sleep(2)
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "Empresa", sheet["A2"] = "Novo Atacarejo"
        sheet["B1"] = "Campanha", sheet["B2"] = campanha_and_data
        workbook.save = ENCARTE_DIR/"campanhas_novoatacarejo.xlsx"
    except:
        print("Não foi possivel salvar")
       
save_in_spreadsheet()
    
driver.quit()
