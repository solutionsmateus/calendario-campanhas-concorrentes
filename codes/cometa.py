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

BASE_URL = "https://cometasupermercados.com.br/ofertas/"

# --- CONFIGURAÇÃO DE CAMINHOS PARA GITHUB ACTIONS ---
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", str(Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Cometa-Supermercados"))
ENCARTE_DIR = Path(OUTPUT_DIR)
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_cometa.xlsx" 


#=== CHROME HEADLESS ===
def iniciar_driver():
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
        

def procurador_campanhas(driver):
    try:
        driver.get(BASE_URL)
        wait = WebDriverWait(driver, 20)
        
        campanhas = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//h3[contains(@class, 'elementor-heading-title')]")
        ))
        
        datas = driver.find_elements(
            By.XPATH, "//div[contains(@class,'jet-listing-dynamic-field__content')]"
        )
        
        textos_datas = [d.text.strip() for d in datas if "Ofertas válidas" in d.text]

        num_flyers = min(len(campanhas), len(textos_datas))
        print(f"Ofertas encontradas: {num_flyers}")

        for i in range(num_flyers):
            dados = {
                "Empresa": "Cometa Supermercados",
                "Campanha_Titulo": campanhas[i].text.strip(),
                "Validade_Texto": textos_datas[i],
                "Cidade": "Fortaleza",
                "Estado": "CE",
                "Data_Coleta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
            
    except Exception as e:
        print(f"Não foi possível processar as campanhas: {e}")

# --- BLOCO PRINCIPAL ---
driver = iniciar_driver()
try:
    procurador_campanhas(driver)
    print("\nProcesso finalizado.")
finally:
    driver.quit()
