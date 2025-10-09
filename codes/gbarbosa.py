import os
import re
import time
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://blog.gbarbosa.com.br/ofertas/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Extraidos-Campanhas/G-Barbosa"
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_gbarbosa.xlsx" 


# ===== Chrome headless =====
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

def save_as_xlsx(data_dict, file_path):
    try:
        df_novo = pd.DataFrame([data_dict])
        if file_path.exists():
            df_existente = pd.read_excel(file_path, engine='openpyxl')
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = df_novo
            
        df_final.to_excel(file_path, index=False, engine='openpyxl')
        print(f"  -> Dados da campanha '{data_dict['Campanha_Titulo']}' salvos para o estado {data_dict['Estado']}.")
    except Exception as e:
        print(f"  -> Erro ao salvar no Excel: {e}")    
        

# A função continua a mesma, recebendo o estado (uf) como parâmetro
def processar_campanhas(uf):
    print(f" Processando campanhas do estado: {uf}")
    time.sleep(2) # Espera para os cards do estado carregarem
    try:
        # Seletores buscam apenas os encartes visíveis do estado selecionado
        campanhas = driver.find_elements(By.XPATH, "//div[@class='encarte-item' and contains(@style,'display: block')]/h2")
        datas = driver.find_elements(By.XPATH, "//div[@class='encarte-item' and contains(@style,'display: block')]/p")
        
        num_flyers = len(campanhas)
        if num_flyers == 0:
            print(f" Nenhuma campanha encontrada para {uf}.")
            return
            
        print(f" {num_flyers} campanhas encontradas para {uf}.")
        
        for i in range(num_flyers):
            titulo_campanha = campanhas[i].text.strip()
            validade_texto = datas[i].text.strip()
            
            dados = {
                'Empresa': 'GBarbosa',
                'Campanha_Titulo': titulo_campanha,
                'Validade_Texto': validade_texto,
                'Cidade': f'Cidades de {uf}',
                'Estado': uf,
                'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
    
    except Exception as e:
        print(f" Não foi possível processar as campanhas de {uf}. Erro: {e}")

driver = build_headless_chrome()
wait = WebDriverWait(driver, 25)

try:
    print("Iniciando processo de coleta no GBarbosa...")
    driver.get(BASE_URL)
    wait.until(EC.presence_of_element_located((By.XPATH, "//button[@class='estado-btn']"))) # Espera os botões carregarem
    
    print("-" * 30)
    processar_campanhas('AL')
    
    print("-" * 30)
    try:
        print("Mudando para o estado SE...")
        botao_estado_se = wait.until(EC.element_to_be_clickable(
            (By.XPATH, f"//button[@class='estado-btn' and text()='SE']")
        ))
        botao_estado_se.click()
        
        processar_campanhas('SE')

    except Exception as e:
        print(f" Erro ao processar o estado SE: {e}")

    print("-" * 30)
    print("\nProcesso finalizado com sucesso.")

except Exception as e:
    print(f"\nOcorreu um erro geral no processo: {e}")
finally:       
    driver.quit()