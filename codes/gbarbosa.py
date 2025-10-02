#Função principal para selecionar e ir com ChromeDrive.

import os
import re
import time
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

#Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://blog.gbarbosa.com.br/ofertas/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Concorrentes/G-Barbosa"
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

# ===== Chrome headless =====
def build_headless_chrome(download_dir: Path):
    options = webdriver.ChromeOptions()
    # Headless moderno e flags de CI
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
    # Preferências de download
    prefs = {
        "download.prompt_for_download": False,
        "download.default_directory": str(download_dir),
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

driver = build_headless_chrome(ENCARTE_DIR)
wait = WebDriverWait(driver, 25)

def processar_campanhas(uf: str):
    print(f"\n Processando campanhas do estado: {uf}")
    driver.get(BASE_URL)
    time.sleep(5)

    try:
        botao_estado = wait.until(EC.element_to_be_clickable((By.XPATH, f'//button[normalize-space()="{"AL"}"]')))
        botao_estado.click()
        time.sleep(3)
        campanha = driver.find_element(By.XPATH, "h3//[contains('text')]")
        campanha = campanha.text
        data = driver.find_element(By.XPATH, "p//[contains('text')]")
        data = data.text
    except:
        print("Estado AL (Alagoas) não processado")
        return
    
    try: 
        botao_estado = wait.until(EC.element_to_be_clickable((By.XPATH, f'//button[normalize-space()="{"SE"}"]')))
        botao_estado.click()
        time.sleep(3)
        campanha = driver.find_element(By.XPATH, "h2//contain('text')")
        campanha.text
        data = driver.find_element(By.XPATH, "p//[contains('text')]")
        data.text
    except:
        print("Estado SE (Sergipe) não processado.")

processar_campanhas()
        
def save_in_spreadsheet(campanha, data):
    try:
        print("Salvando na planilha")
        time.sleep(2)
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "Empresa", sheet["A2"] = "GBarbosa"
        sheet["B1"] = "Campanha", sheet["B2"] = campanha
        sheet["C1"] = "Data", sheet["C2"] = data
        workbook.save = ENCARTE_DIR/"campanhas_gbarbosa.xlsx"
    except:
        print("Não foi possivel salvar")
       
save_in_spreadsheet()
    
driver.quit()
print("\nFinalizado.")
