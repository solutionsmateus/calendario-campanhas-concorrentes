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
import pandas as pd
from openpyxl import Workbook
from datetime import datetime


XLSX_FILE_PATH = Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Atakarejo"
XLSX_FILE_PATH.mkdir(parents=True, exist_ok=True)

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
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)

driver = build_headless_chrome()
driver.get("https://atakarejo.com.br/cidade/vitoria-da-conquista")

links = driver.find_elements(By.XPATH, '//a[contains(@class, "button-download-ofertas")]')
print(f"{len(links)} encarte(s) encontrado(s).")

def encontrar_data():
    try: 
        enc_data = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, '//h3[contains("TEXT")]')))
    except:
        return "sem_data"
    
    for div in enc_data:
        texto = div.text.strip()
        if texto:
            nome_pasta = re.sub(r'[\\/*?:"<>|\s]', '_', texto)
            return nome_pasta
    return "sem_data"

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
        print(f" Dados anexados/salvos em {file_path.name}")
    except Exception as e:
        print(f" Erro ao salvar no Excel: {e}")    
        
        
    
#Procurar campanhas na Pagina HTML
def procurar_campanhas():
    print("Selecionando a campanha, data, mes e dia")
    campanha_and_data = driver.find_element(By.XPATH, "//h3[contains('text')]")
    min(len(campanha_and_data))
    
    try:
        print(f"Lido encarte {campanha_and_data}" )
        campanha_and_data.text
        dados = {
                'Empresa': 'Atakarejo',
                'Campanha_Titulo': campanha_and_data,  # Texto limpo
                'Validade_Texto': campanha_and_data,  # Texto limpo
                'Cidade': 'Vitória da Conquista',
                'Estado': 'Bahia',
                'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
        save_as_xlsx(dados, XLSX_FILE_PATH)
        
    except:
        print("Campanhas e datas não encontradas")
    

"""for i, link in enumerate(links):
    url_pdf = link.get_attribute("href")
    if not url_pdf:
        continue
    nome = f"encarte_{i+1}.pdf"
    caminho = ENCARTE_DIR / nome
    try:
        resp = requests.get(url_pdf, timeout=20)
        if resp.status_code == 200:
            with open(caminho, "wb") as f:
                f.write(resp.content)
            print(f"Baixado: {caminho.name}")
        else:
            print(f"Falha no download ({resp.status_code}): {url_pdf}")
    except Exception as e:
        print(f"Erro ao baixar {url_pdf}: {e}")"""""

driver.quit()
