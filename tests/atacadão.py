import os
import re
import time
import requests
import pandas as pd
from openpyxl import Workbook
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime

 

LOJAS_ESTADOS = {
    "MA": ("São Luís", "São Luís"),
    "AL": ("Maceió", "Maceió Praia"),
    "CE": ("Fortaleza", "Fortaleza Fátima"),
    "PA": ("Belém", "Belém Portal da Amazônia"),
    "PB": ("João Pessoa", "João Pessoa Bessa"),
    "PE": ("Recife", "Recife Avenida Recife"),
    "PI": ("Teresina", "Teresina Primavera"),
    "SE": ("Aracaju", "Aracaju Tancredo Neves"),
    "BA": ("Vitória Da Conquista", "Vitória da Conquista"),
    "MA": ("São Luís", "São Luís"),

}

ENCARTE_DIR = Path.home() / "Desktop/Encartes-Concorrentes/Atacadão"

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_atacadão.xlsx" 

BASE_URL = "https://www.atacadao.com.br/institucional/nossas-lojas"

# === CHROME HEADLESS ===
def build_headless_chrome():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")              
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--window-size=1920,1080")     # substitui start-maximized
    options.add_argument("--lang=pt-BR,pt")
    options.add_argument("--start-maximized") 
    # options.add_argument("--enable-unsafe-swiftshader")
    user_agent = os.environ.get("HTTP_UA", "Mozilla/5.0 (...)") 
    options.add_argument(f"--user-agent={user_agent}")

    return webdriver.Chrome(options=options)

driver = build_headless_chrome()
wait = WebDriverWait(driver, 30)

def encontrar_data():
    try:
        enc_data = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//p[contains(@class, "text-xs text-neutral-400")]')))
    except:
        return "sem_data"
    
    for div in enc_data:
        texto = div.text.strip()
        if texto:
            nome_pasta = re.sub(r'[\\/*?:"<>|\s]', '_', texto)
            return nome_pasta
    return "sem_data"

def clicar_confirmar():
    try:
        confirmar_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Confirmar']")))
        confirmar_button.click()
    except:
        pass

def selecionar_uf_cidade(uf, cidade):
    Select(
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[contains(@class, 'md:w-[96px]')]")
        ))
    ).select_by_value(uf)
    time.sleep(1)
    Select(
        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//select[contains(@class, 'md:w-[360px]')]")
        ))
    ).select_by_visible_text(cidade)
    time.sleep(1)

def clicar_loja_por_nome(loja_nome):
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-testid='store-card']")))
    lojas = driver.find_elements(By.CSS_SELECTOR, "[data-testid='store-card']")
    for loja in lojas:
        try:
            titulo = loja.find_element(By.TAG_NAME, "h1").text
            if loja_nome.lower() in titulo.lower():
                botao = loja.find_element(By.TAG_NAME, "a")
                print(f"Acessando loja: {titulo}")
                botao.click()
                return titulo
        except:
            continue
    print(f" Loja '{loja_nome}' não encontrada.")
    return None

#Select the elements and save spreadsheet       
def save_as_xlsx(data_dict, file_path):
    try:
        df_novo = pd.DataFrame([data_dict])
        if file_path.exists():
            df_existente = pd.read_excel(file_path, engine='openpyxl')
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = df_novo
            
        df_final.to_excel(file_path, index=False, engine='openpyxl')
        print(f"Dados anexados/salvos em {file_path}")
    except Exception as e:
        print(f"Erro ao salvar no Excel: {e}")


def processar_campanhas(uf, cidade, loja_nome):
    print("Selecionando a campanha, data, mes e dia")
    wait_local = WebDriverWait(driver, 15) 

    try:
        campanhas_elementos = wait_local.until(
            EC.presence_of_all_elements_located((By.XPATH, "//h1[contains(@class, 'text-sm font-bold text-atc-primary')]"))
        )
        datas_elementos = wait_local.until(
            EC.presence_of_all_elements_located((By.XPATH, "//p[contains(@class , 'text-xs text-neutral-400')]"))
        )

        for jornal_num, (campanha_elem, data_elem) in enumerate(zip(campanhas_elementos, datas_elementos), start=1):
            
            campanha_texto = campanha_elem.get_attribute('outerHTML') 
            data_texto_item = data_elem.get_attribute('outerHTML')
            
            if not campanha_texto:
                print(f"Aviso: Campanha {jornal_num} não possui texto.")
                continue
                
            dados = {
                'Empresa': 'Atacadão',
                'Campanha': campanha_texto,
                'Cidade': cidade,
                'Estado': uf,
                'Loja': loja_nome,
                'Jornal Número': jornal_num,
                'Validade Texto': data_texto_item, 
                'Data Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
    
    except Exception as e:
        print(f"Não foi possível processar as campanhas. Erro: {e}")
    

try:
    driver.get(BASE_URL)
    clicar_confirmar()

    for uf, (cidade, loja_nome) in LOJAS_ESTADOS.items():
        print(f"\n Estado: {uf} | Cidade: {cidade} | Loja: {loja_nome}")
        driver.get(BASE_URL)
        time.sleep(2)
        clicar_confirmar()

        selecionar_uf_cidade(uf, cidade)
        nome_loja_encontrada = clicar_loja_por_nome(loja_nome)

        if nome_loja_encontrada:
            processar_campanhas(uf, cidade, nome_loja_encontrada, 1)
            time.sleep(1)

except Exception as e:
    print(f" Erro geral: {e}")

finally:
    print(" Execução finalizada")
    driver.quit()
