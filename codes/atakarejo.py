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

# --- CONFIGURAÇÃO DE CAMINHOS PARA GITHUB ACTIONS ---
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", str(Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Atakarejo"))
ENCARTE_DIR = Path(OUTPUT_DIR)
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_atakarejo.xlsx" 

# === CHROME HEADLESS ===
def build_headless_chrome():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")               
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--window-size=1920,1080")     
    options.add_argument("--lang=pt-BR,pt")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/5.37.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)

#Save as XLSX
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
        
#Procurar campanhas na Pagina HTML
def procurar_campanhas(driver):
    try:
        cards = driver.find_elements(By.CSS_SELECTOR, "div.subitem.active")
        print(f"Encontrados: {len(cards)} encartes.")
        if not cards:
            print("Nenhum card de campanha encontrado.")
            return

        for i, card in enumerate(cards, 1):
            h3 = card.find_elements(By.TAG_NAME, "h3")
            h4 = card.find_elements(By.TAG_NAME, "h4")

            dados = {
                "Empresa": "Atakarejo",
                "Campanha_Titulo": h3[0].text.strip() if len(h3) > 0 else "Título não encontrado",
                "Validade_Texto": (h3[1].text.strip() if len(h3) > 1
                                   else (h4[0].text.strip() if h4 else "Validade não encontrada")),
                "Cidade": "Vitória da Conquista",
                "Estado": "BA",
                "Data_Coleta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)

    except Exception as e:
        print(f"Erro ao procurar campanhas/datas: {e}")

# --- BLOCO PRINCIPAL ---
driver = build_headless_chrome()
try:
    driver.get("https://atakarejo.com.br/cidade/vitoria-da-conquista")
    time.sleep(3) # Espera a página carregar
    procurar_campanhas(driver)
    print("\nProcesso finalizado.")
except Exception as e:
    print(f"Ocorreu um erro geral: {e}")
finally:
    driver.quit()
