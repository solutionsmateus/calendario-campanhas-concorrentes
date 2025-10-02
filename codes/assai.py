#Função principal para selecionar e ir com ChromeDrive.

import os
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import re


#Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook


# CHANGE: Ajuste do mapeamento da Bahia para o NOME DA LOJA (e não a região)
LOJAS_ESTADOS = {
    "Maranhão": "Assaí Angelim",
    "Alagoas": "Assaí Maceió Farol",
    "Ceará": "Assaí Bezerra M (Fortaleza)",
    "Pará": "Assaí Belém",
    "Paraíba": "Assaí João Pessoa Geisel",
    "Pernambuco": "Assaí Avenida Recife",
    "Piauí": "Assaí Teresina",
    "Sergipe": "Assaí Aracaju",
    "Bahia": "Assaí Vitória da Conquista",  # CHANGE
}

# CHANGE: Região preferida por estado (usado quando existe select.regiao)
REGIAO_POR_ESTADO = {
    "Bahia": "Interior",  
}

BASE_URL = "https://www.assai.com.br/ofertas"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Concorrentes/Assai"

# === HEADLESS CHROME ===
def build_headless_chrome():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--window-size=1920,1080")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    options.add_argument("--lang=pt-BR,pt")
    return webdriver.Chrome(options=options)

driver = build_headless_chrome()
wait = WebDriverWait(driver, 30)

def encontrar_data():
    try:
        enc_data = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, '//div[contains(@class, "ofertas-tab-validade")]'))
        )
    except:
        return "sem_data"
    
    for div in enc_data:
        texto = div.text.strip()
        if texto:
            nome_pasta = re.sub(r'[\\/*?:"<>|\s]', '_', texto)
            return nome_pasta
    return "sem_data"

def aguardar_elemento(seletor, by=By.CSS_SELECTOR, timeout=15):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, seletor)))

def clicar_elemento(seletor, by=By.CSS_SELECTOR):
    element = wait.until(EC.element_to_be_clickable((by, seletor)))
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.5)
    element.click()

def scroll_down_and_up():
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight/3);")
    time.sleep(0.5)
    driver.execute_script("window.scrollTo(0, 1);")
    time.sleep(0.5)

def processar_campanhas(jornal_num):
    page_num = 1
    while True:
        print(f"Processando a campanha do  {page_num} do jornal {jornal_num}...")
        data = driver.find_element((By.XPATH, "//div[contains(@class, 'ofertas-tab-validade')]"))
        data.text
        Select(data)

        try:
            next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.slick-next")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
            time.sleep(0.5)
            next_button.click()
            time.sleep(2)
            page_num += 1
        except:
            break
        
def save_in_spreadsheet(data):
    try:
        print("Salvando na planilha")
        time.sleep(2)
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "Empresa", sheet["A2"] = "Assai"
        sheet["B1"] = "Campanha", sheet["B2"] = "Campanha"
        sheet["B1"] = "Campanha", sheet["C2"] = data
        workbook.save = ENCARTE_DIR/"campanhas_novoatacarejo.xlsx"
    except:
        print("Não foi possivel salvar")
       

# CHANGE: helper para selecionar por "contém texto" (robusto contra variações de acento/espaço)
def select_by_visible_text_contains(select_el, target_text, timeout=10):
    WebDriverWait(driver, timeout).until(lambda d: len(select_el.find_elements(By.TAG_NAME, "option")) > 0)
    sel = Select(select_el)
    opts = select_el.find_elements(By.TAG_NAME, "option")
    alvo_norm = target_text.strip().lower()
    for o in opts:
        if alvo_norm in o.text.strip().lower():
            sel.select_by_visible_text(o.text)
            return True
    return False

try:
    driver.get(BASE_URL)
    time.sleep(2)

    try:
        clicar_elemento("button.ot-close-icon")
    except:
        pass

    clicar_elemento("a.seletor-loja")
    time.sleep(1)

    for estado, loja in LOJAS_ESTADOS.items():
        print(f" Processando: {estado} - {loja}")

        estado_select = aguardar_elemento("select.estado")
        Select(estado_select).select_by_visible_text(estado)
        time.sleep(1)
        if estado in REGIAO_POR_ESTADO:
            try:
                regiao_select_element = aguardar_elemento("select.regiao", timeout=15)
                Select(regiao_select_element).select_by_visible_text(REGIAO_POR_ESTADO[estado])
                aguardar_elemento("select.loja option[value]", timeout=20)
                time.sleep(0.5)
            except Exception as e:
                print(f" Não foi possível selecionar a região para {estado}: {e}")
        loja_select = aguardar_elemento("select.loja", timeout=20)
        try:
            Select(loja_select).select_by_visible_text(loja)
        except:
            ok = select_by_visible_text_contains(loja_select, loja)
            if not ok:
                raise RuntimeError(f"Não encontrei a loja '{loja}' no estado {estado}")

        time.sleep(0.8)

        clicar_elemento("button.confirmar")
        time.sleep(1)
        aguardar_elemento("div.ofertas-slider", timeout=30)
        data_nome = encontrar_data()
        nome_loja = loja.replace(' ', '_').replace('(', '').replace(')', '')
        scroll_down_and_up()
        processar_campanhas(1, ENCARTE_DIR)

        for i in range(2, 4):
            try:
                clicar_elemento(f"//button[contains(., 'Jornal de Ofertas {i}')]", By.XPATH)
                time.sleep(3)
                aguardar_elemento("div.ofertas-slider", timeout=30)
                scroll_down_and_up()
                processar_campanhas(i, ENCARTE_DIR)
            except Exception as e:
                print(f" Jornal {i} indisponível para {loja}: {str(e)}")

        clicar_elemento("a.seletor-loja")
        time.sleep(2)

    print("Todos os encartes foram processados!")

except Exception as e:
    print(f"Erro crítico: {str(e)}")
    try:
        driver.save_screenshot(str(ENCARTE_DIR / "erro_encartes.png"))
    except Exception as _:
        pass
finally:
    driver.quit()
