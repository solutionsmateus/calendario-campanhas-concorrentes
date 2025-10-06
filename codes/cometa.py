Função principal para selecionar e ir com ChromeDrive.

import os
import re
import time
import requests
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

#Importar bibliotecas para nova função de procurar campanhas.
import pandas as pd
from openpyxl import Workbook

BASE_URL = "https://cometasupermercados.com.br/ofertas/"
ENCARTE_DIR = Path.home() / "Desktop/Encartes-Concorrentes/Cometa-Supermercados"
ENCARTE_DIR.mkdir(parents=True, exist_ok=True)

#=== CHROME HEADLESS ===
def iniciar_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")               #headless moderno
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-features=VizDisplayCompositor")
    options.add_argument("--window-size=1920,1080")      #substitui start-maximized no headless
    options.add_argument("--lang=pt-BR,pt")
    options.add_argument(
        "--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
    return webdriver.Chrome(options=options)

def procurador_campanhas():
    try:
       print("Procurando campanhas")
       driver = iniciar_driver()
       wait = WebDriverWait(driver, 10)
       driver.get(BASE_URL)
       time.sleep(3)
       campanha = driver.find_elements(By.XPATH, "//h3[contains(@class, 'elementor-heading-title elementor-size-default')]")
       campanha = campanha.text
       Select(campanha)
       data = driver.find_elements(By.XPATH, "div//[contains(@class, 'jet-listing-dynamic-field__content')]")
       data = data.text
       Select(data)
    except:
        print("Não foi possivel processar os encartes.")

procurador_campanhas()

# Select the elements and save in spreeadsheet       
def save_in_spreadsheet(campanha, data):
    try:
        print("Salvando na planilha")
        time.sleep(2)
        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"] = "Empresa", sheet["A2"] = "Cometa Supermercados"
        sheet["B1"] = "Campanha", sheet["B2"] = campanha
        sheet["C1"] = "Data", sheet["C3"] = data
        workbook.save = ENCARTE_DIR/"campanhas_cometa.xlsx"
    except:
        print("Não foi possivel salvar")

save_in_spreadsheet()
