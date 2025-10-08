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
from datetime import datetime

#Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://cometasupermercados.com.br/ofertas/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Cometa-Supermercados"
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
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
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
        print(f" Dados anexados/salvos em {file_path.name}")
    except Exception as e:
        print(f" Erro ao salvar no Excel: {e}")    
        

def procurador_campanhas():
    try:
        driver.get(BASE_URL)
        wait = WebDriverWait(driver, 20)
        campanhas = wait.until(EC.presence_of_all_elements_located(
            (By.XPATH, "//h3[contains(@class,'jet-listing-dynamic-field__content')]")
        ))
        datas = driver.find_elements(
            By.XPATH, "//h4[contains(@class,'jet-listing-dynamic-field__content')]"
        )

        num_flyers = min(len(campanhas), len(datas))
        print(f"Ofertas encontradas: {num_flyers}")

        for i, (c_elem, d_elem) in enumerate(zip(campanhas, datas), start=1):
            dados = {
                "Empresa": "Cometa Supermercados",
                "Campanha_Titulo": c_elem.text.strip(),
                "Validade_Texto": d_elem.text.strip(),
                "Cidade": "Fortaleza",
                "Estado": "CE",
                "Data_Coleta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
            print(f"Campanha {i} salva.")
    except Exception as e:
        print(f"Não foi possível processar as campanhas: {e}")

        
driver = iniciar_driver()
try:
    procurador_campanhas()
finally:
    driver.quit()