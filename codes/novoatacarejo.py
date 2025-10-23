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

# Esta configuração já estava correta para o GitHub Actions
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", str(Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Novo-Atacarejo"))
ENCARTE_DIR = Path(OUTPUT_DIR)
ENCARTE_DIR.mkdir(parents=True, exist_ok=True) # Garante que o diretório existe

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
        
def procurar_campanhas(driver, wait, cidade_selecionada):
    print(f"Processando campanhas para {cidade_selecionada}")
    try:
        # Espera os cards de tabloide carregarem
        wait.until(EC.presence_of_element_located((By.XPATH, "//div[@class='card-tabloids']")))
        
        cards_tabloide = driver.find_elements(By.XPATH, "//div[@class='card-tabloids']")
        print(f"{len(cards_tabloide)} campanhas encontradas.")

        for card in cards_tabloide:
            try:
                titulo = card.find_element(By.TAG_NAME, 'h3').text.strip()
                validade = card.find_element(By.TAG_NAME, 'h6').text.strip()
                
                dados = {
                    'Empresa': 'Novo Atacarejo',
                    'Campanha_Titulo': titulo,
                    'Validade_Texto': validade,
                    'Cidade': cidade_selecionada,
                    'Estado': 'PE', # Site parece focado em PE
                    'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                save_as_xlsx(dados, XLSX_FILE_PATH)
            except Exception as e:
                print(f"  -> Erro ao processar um card de campanha: {e}")

    except Exception as e:
        print(f"Não foi possivel processar as campanhas para {cidade_selecionada}. Erro: {e}")

# --- BLOCO PRINCIPAL ---
driver = build_headless_chrome()
wait = WebDriverWait(driver, 20)
try:
    driver.get(BASE_URL)
    cidade = "Olinda"
    
    print(f"Selecionando loja: {cidade}")
    select_element = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "select.select")))
    Select(select_element).select_by_visible_text(cidade)
    time.sleep(4)  # Aguarda carregar os tabloides

    procurar_campanhas(driver, wait, cidade)
    print("\nProcesso finalizado.")
except Exception as e:
    print(f"Ocorreu um erro geral: {e}")
finally:
    driver.quit()
