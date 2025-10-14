import os
import re
import time
import requests
import pandas as pd
from openpyxl import Workbook
import urllib.parse
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime


LOJAS_ESTADOS = {
    "BA": ("Vitória Da Conquista", "Vitória da Conquista"),
    "AL": ("Maceió", "Maceió Praia"),
    "CE": ("Fortaleza", "Fortaleza Fátima"),
    "PA": ("Belém", "Belém Portal da Amazônia"),
    "PB": ("João Pessoa", "João Pessoa Bessa"),
    "PE": ("Recife", "Recife Avenida Recife"),
    "PI": ("Teresina", "Teresina Primavera"),
    "SE": ("Aracaju", "Aracaju Tancredo Neves"),
    "MA": ("São Luís", "São Luís"),
}


# Esta configuração já estava correta para o GitHub Actions
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", str(Path.home() / "Desktop/Encartes-Extraidos-Campanhas/Atacadão"))
ENCARTE_DIR = Path(OUTPUT_DIR)
ENCARTE_DIR.mkdir(parents=True, exist_ok=True) # Garante que o diretório existe

XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_atacadao.xlsx" 

BASE_URL = "https://www.atacadao.com.br/institucional/nossas-lojas"


driver = None
wait = None

#Buiding Headless Mode for Chrome
def build_headless_chrome():
    """Configura e retorna o driver Chrome em modo Headless."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--lang=pt-BR,pt")
    options.add_argument("--start-maximized") 
    user_agent = os.environ.get("HTTP_UA", "Mozilla/5.0 (...)") # Usa a ENV do YAML
    options.add_argument(f"--user-agent={user_agent}")

    global driver, wait
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    return driver

#Functions to process Flyers
def encontrar_data():
    """Encontra a data do encarte para criar o nome da pasta (não usada na versão final, mas mantida)."""
    try:
        enc_data = wait.until(
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
    """Tenta clicar no botão 'Confirmar' para fechar pop-ups."""
    try:
        confirmar_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Confirmar']")))
        confirmar_button.click()
    except:
        pass

def selecionar_uf_cidade(uf, cidade):
    """Seleciona o Estado (UF) e a Cidade nos dropdowns."""
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
    """Localiza e clica no link da loja com base no nome parcial."""
    wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "[data-testid='store-card']")))
    lojas = driver.find_elements(By.CSS_SELECTOR, "[data-testid='store-card']")
    for loja in lojas:
        try:
            titulo = loja.find_element(By.TAG_NAME, "h1").text
            if loja_nome.lower() in titulo.lower():
                botao = loja.find_element(By.TAG_NAME, "a")
                print(f"Acessando loja: {titulo}")
                botao.click()
                return titulo # Retorna o nome completo da loja
        except:
            continue
    print(f" Loja '{loja_nome}' não encontrada.")
    return None

#Function to save as xlsx
def save_as_xlsx(data_dict, file_path):
    """Carrega dados existentes, anexa os novos dados e salva a planilha."""
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


#Main Function
def processar_campanhas(uf, cidade, loja_nome):
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, "//h1[contains(@class, 'text-sm font-bold text-atc-primary')]"))
        )

        campanhas = driver.find_elements(By.XPATH, "//h1[contains(@class, 'text-sm font-bold text-atc-primary')]")
        data = driver.find_elements(By.XPATH, "//p[contains(@class , 'text-xs text-neutral-400')]")
        
        num_flyers = min(len(campanhas), len(data))
        print(f" Encontrados {num_flyers} encartes para extração.")

        for i in range(num_flyers):
            
            campanha_titulo = campanhas[i].text 
            data_validade = data[i].text
            
            jornal_num = i 
            
            dados = {
                'Empresa': 'Atacadão',
                'Campanha_Titulo': campanha_titulo,  # Texto limpo
                'Validade_Texto': data_validade,  # Texto limpo
                'Campanha_HTML_Bruto': campanhas[i].get_attribute('outerHTML'), # Se precisar do HTML
                'Validade_HTML_Bruto': data[i].get_attribute('outerHTML'),     # Se precisar do HTML
                'Cidade': cidade,
                'Estado': uf,
                'Loja': loja_nome,
                'Jornal_Número': jornal_num,
                'Data_Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            save_as_xlsx(dados, XLSX_FILE_PATH)
        
    except Exception as e:
        print(f" Não foi possível processar as campanhas da loja {loja_nome}. Erro: {e}")


build_headless_chrome() 

try:
    print("Procurando os Encartes")
    
    driver.get(BASE_URL)
    clicar_confirmar()
    time.sleep(2) 

    for uf, (cidade, loja_nome_busca) in LOJAS_ESTADOS.items():
        print(f" Processando UF: {uf} | Cidade: {cidade} | Loja: {loja_nome_busca} ")
        
        driver.get(BASE_URL)
        clicar_confirmar() 
        
        selecionar_uf_cidade(uf, cidade)
        
        nome_loja_encontrada = clicar_loja_por_nome(loja_nome_busca)

        if nome_loja_encontrada:
            processar_campanhas(uf, cidade, nome_loja_encontrada)
        
except Exception as e:
    print(f" Erro geral no script: {e}")

finally:
    print("\nExecução finalizada.")
    driver.quit()
