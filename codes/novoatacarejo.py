# Função principal para selecionar e ir com ChromeDrive.

import os
import re
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://novoatacarejo.com/oferta/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Novo-Atacarejo"

ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_novoatacarejo.xlsx" 

def build_headless_chrome():
    options = webdriver.ChromeOptions()
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
wait = WebDriverWait(driver, 20)

def encontrar_data():
  #   "h6 - TEXT LOCATION OF DATES IN PAGE"
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

def save_as_xlsx(data_dict, file_path):
    try:
        df_novo = pd.DataFrame([data_dict])
        if file_path.exists():
            df_existente = pd.read_excel(file_path, engine='openpyxl')
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = df_novo
            
        df_final.to_excel(file_path, index=False, engine='openpyxl')
        print(f" Dados anexados/salvos em {file_path.name}")
    except Exception as e:
        print(f" Erro ao salvar no Excel: {e}")    
        

def procurar_campanhas():
    print("Processando campanhas")
    
    try:
        campanha_and_data = driver.find_elements(By.XPATH, "h6//contains('text')")
        num_flyers = min(len(campanha_and_data))
        for i in range(num_flyers):
            campanha_and_data[i].text
            
            dados = {
                'Empresa': 'GBarbosa',
                'Campanha_Titulo': campanha_and_data,  # Texto limpo
                'Validade_Texto': campanha_and_data,  # Texto limpo
                'Cidade': 'Maceió e Aracaju',
                'Estado': 'AL e SE',
                'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
    except:
        print(f"Não foi possivel processar as campanhas {num_flyers}")
  
driver.quit()
