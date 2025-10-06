#Função principal para selecionar e ir com ChromeDrive.

import os
import re
import time
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

#Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://blog.gbarbosa.com.br/ofertas/"
XLSX_FILE_PATH = Path.home() / "Desktop/Encartes-Extraidos-Campanhas/G-Barbosa"
XLSX_FILE_PATH.mkdir(parents=True, exist_ok=True)

# ===== Chrome headless =====
def build_headless_chrome(download_dir: Path):
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
#     Preferências de download
    prefs = {
        "download.prompt_for_download": False,
        "download.default_directory": str(download_dir),
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return webdriver.Chrome(options=options)

driver = build_headless_chrome(XLSX_FILE_PATH)
wait = WebDriverWait(driver, 25)

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
        

def processar_campanhas(uf: str):
    print(f"\n Processando campanhas do estado: {uf}")
    time.sleep(5)
    try:
        campanha = driver.find_elements(By.XPATH, "h3//[contains('text')]")
        data = driver.find_elements(By.XPATH, "p//[contains('text')]")
        
        num_flyers = min(len(campanha), len(data))
        
        print(f"Encartes encontrados {num_flyers}")
        
        for i in range(num_flyers):
            campanha[i].text
            data[i].text
            
            dados = {
                'Empresa': 'GBarbosa',
                'Campanha_Titulo': campanha,  # Texto limpo
                'Validade_Texto': data,  # Texto limpo
                'Cidade': 'Maceió e Aracaju',
                'Estado': 'AL e SE',
                'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
    
    except:
        print(f"Não foi possivel processar as campanhas {num_flyers}")
        
driver.quit()