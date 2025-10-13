import os
import re
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
from openpyxl import Workbook
from datetime import datetime

BASE_URL = "https://frangolandia.com/encartes/"

# --- CONFIGURAÇÃO DE CAMINHOS PARA GITHUB ACTIONS ---
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", str(Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Frangolandia"))
ENCARTE_DIR = Path(OUTPUT_DIR)
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_frangolandia.xlsx" 

def build_headless_chrome():
    options = webdriver.ChromeOptions()
    prefs = {
        "download.prompt_for_download": False,
        "plugins.always_open_pdf_externally": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--lang=pt-BR,pt")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/5.37.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)

def save_as_xlsx(data_dict, file_path):
    try:
        df_novo = pd.DataFrame([data_dict])
        if file_path.exists():
            df_existente = pd.read_excel(file_path, engine='openpyxl')
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = df_novo
            
        df_final.to_excel(file_path, index=False, engine='openpyxl')
        print(f" Dados da campanha '{data_dict['Campanha_Titulo']}' salvos.")
    except Exception as e:
        print(f" Erro ao salvar no Excel: {e}")   
        
def procurar_campanhas(driver):
    try:
        datas = driver.find_elements(By.XPATH, "//span[contains(@class, 'elementor-button-text')]")
        campanhas = driver.find_elements(By.XPATH, "//h3[contains(@class, 'elementor-heading-title')]")
        
        campanhas_filtradas = [c for c in campanhas if c.text.strip() != ""]
        
        num_flyers = min(len(campanhas_filtradas), len(datas))
        print(f"Campanhas encontradas: {num_flyers}")
        
        for i in range(num_flyers):
           titulo_campanha = campanhas_filtradas[i].text.strip()
           data_titulo = datas[i].text.strip()   
           
           dados = {
                'Empresa': 'Frangolandia',
                'Campanha_Titulo': titulo_campanha,
                'Validade_Texto': data_titulo,
                'Cidade': 'Fortaleza',
                'Estado': 'CE',
                'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
           save_as_xlsx(dados, XLSX_FILE_PATH)

    except Exception as e:
        print(f"Não foi possivel processar as campanhas. Erro: {e}")
                
# --- BLOCO PRINCIPAL ---       
driver = build_headless_chrome()
try:
    driver.get(BASE_URL)
    wait = WebDriverWait(driver, 15)
    procurar_campanhas(driver)
    print("\nProcesso finalizado.")
except Exception as e:
    print(f"Erro geral: {e}")
finally:
    driver.quit()
