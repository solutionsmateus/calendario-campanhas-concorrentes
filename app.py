import os
import re
import time
import pytesseract
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

#Initial Configuration to run in paralelal processesors

#DIR - OF DOCUMENT .DOCX
ENCARTE_DIR = "Desktop/Campanhas_get\Campanhas_doc.docx"
#DIR - DOS ENCARTES - ASSAI PARA OCR
ENCARTE_DIR_ENC_ASSAI = "Desktop/Encartes-Concorrentes/Assai"
#DIR - DOS ENCARTES - GBARBOSA PARA OCR 
ENCARTE_DIR_ENC_GBARBOSA = "Desktop/Encartes-Concorrentes/GBarbosa"


BASE_URLS = {
    "Assaí": "https://www.assai.com.br/ofertas",
    "Atacadão": "https://www.atacadao.com.br/institucional/nossas-lojas",
    "Cometa_Supermercados": "https://cometasupermercados.com.br/ofertas/",
    "Atakarejo": "https://atakarejo.com.br/cidade/vitoria-da-conquista",
    "Frangolandia": "https://frangolandia.com/encartes/",
    "GBarbosa": "https://blog.gbarbosa.com.br/ofertas/",
    "Novo_Atacarejo": "https://novoatacarejo.com/oferta/"
}

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = webdriver(driver, 30)

#Atacadão.LOJAS - FIND

def config_atacadão():
    LOJAS = {
        "MA": ("São Luís", "São Luís"),
        "AL": ("Maceió", "Maceió Praia"),
        "CE": ("Fortaleza", "Fortaleza Fátima"),
        "PA": ("Belém", "Belém Portal da Amazônia"),
        "PB": ("João Pessoa", "João Pessoa Bessa"),
        "PE": ("Recife", "Recife Avenida Recife"),
        "PI": ("Teresina", "Teresina Primavera"),
        "SE": ("Aracaju", "Aracaju Tancredo Neves"),
        "BA": ("Vitória Da Conquista", "Vitória da Conquista"),
    }
    
#Assai.LOJAS - FIND

def config_assai():
    LOJA_ESTADO = {
        "Maranhão": "Assaí Angelim",
        "Alagoas": "Assaí Maceió Farol",
        "Ceará": "Assaí Bezerra M (Fortaleza)",
        "Pará": "Assaí Belém",
        "Paraíba": "Assaí João Pessoa Geisel",
        "Pernambuco": "Assaí Avenida Recife",
        "Piauí": "Assaí Teresina",
        "Sergipe": "Assaí Aracaju",
        "Bahia": "Interior Vitória da Conquista",
    }


#Switch handle windows of all pages

for url in BASE_URLS:
     driver.get(url).wait(3)
     driver.switch_to.window(driver.window_handles[0])
     
     
#Open document Campanhas_doc.docx in computer
driver_path_open = os.path.abspath('Desktop/Campanhas_get/Campanhas_doc.docx')
driver_file_open = f'file:///Desktop/Campanhas_get/Campanhas_doc.docx'
     
#Class Atacadão

class Atacadão:
    
    def init_atacadão():
        driver.get(BASE_URLS["Atacadão"])
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
            EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Confirmar']"))
            )
            confirmar_button.click()
        except:
            pass  
    
    def selecionar_uf_cidade(uf, cidade):
        Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'md:w-[96px]')]")))).select_by_value(uf)
        time.sleep(1)
        Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@class, 'md:w-[360px]')]")))).select_by_visible_text(cidade)
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
    
    def find_text_camp_data():
        #Data e Campanha de Cada Encarte do Assai
        for loja in config_atacadão['LOJAS']:
            camp_with_data_enc = driver.find_element(By.XPATH, '//h1[contains(@class, "text-sm font-bold text-atc-primary")]')
            camp_with_data_enc.split('\n')
            camp_with_data_enc.select()
        
#Class Assaí
            
class Assaí:
      driver.get(BASE_URLS["Atakarejo"])
      #FOR MAKE OCR OF IMAGE IN PATH["Desktop/Encartes-Concorrentes"] AND SAVE ON THE DOCUMENT.     
      def enc_camp():
            data_enc = driver.find_element(By.XPATH, 'div//[contains(@class, "ofertas-tab-validade"]')
            data_enc.split('\n')
            data_enc.select() 
            try:
                ENCARTE_DIR_ENC_ASSAI(len[0-7])
                pdfs =[]
                images = ENCARTE_DIR(pdfs)
                for i, img in enumerate(len(images)):
                    text = pytesseract.image_to_string(img)
                    full_text += f"\n--- Page {i+1} ---\n{text}"
            except:
                print("Not Detected")
        

#Class NovoAtacarejo
class NovoAtacarejo:
    driver.get(BASE_URLS["Atakarejo"])
    
    def enc_camp():
        #Encontrar data - Novo Atacarejo
        try:
            data_enc = driver.find_element(By.XPATH, 'h6//[contains("TEXT")]')
            data_enc.split('\n')
            data_enc.select()
        except:
            return "Nothing located"

#Class Cometa 
class Cometa:
    driver.get(BASE_URLS["Atakarejo"])
    
    def enc_camp():
        try:
            #Data da Campanha - Cometa Supermercados
            data_enc = driver.find_element(By.XPATH 'div//[contains(@class, "jet-listing-dynamic-field__content"])')
            data_enc.split('\n')
            data_enc.select()
            
            #Nome da Campanha - Cometa Supermercados
            
            camp_enc = driver.find_element(By.XPATH, 'h3//[contains(@class, "elementor-heading-title elementor-size-default")]')
            camp_enc.split('\n')
            camp_enc.select()
            
        except:
            return "Nothing located"
        

#Class GBarbosa    
class GBarbosa:
    driver.get(BASE_URLS["Gbarbosa"])
    
    def enc_camp():
        try:
            enc_camp = driver.find_element(By.XPATH, '')
        except:
            return "Nothing located"

#Class Frangolandia   
class Frangolandia:
    driver.get(BASE_URLS["Frangolandia"])
    
    def enc_camp():
        try:
            #Nome da Campanha do Frangolandia
            nome_enc = driver.find_element(By.XPATH, '//h3[contains("elementor-heading-title elementor-size-default")]')
            nome_enc.split('\n')
            nome_enc.select()
            
            #Data da Campanha do Frangolandia
            data_enc = driver.find_element(By.XPATH, '//span[contains(@class,"elementor-button-text"]')
            data_enc.split('\n')
            data_enc.select()
        
        except:
            return "Nothing finded"

#Class Atakarejo
class Atakarejo:
    driver.get(BASE_URLS["Atakarejo"])
    
    def enc_camp():
        try:
            #Data e Campanha do Atakarejo
            enc_camp_text = driver.find_element(By.XPATH, '//h3[contains("TEXT")]')
            enc_camp_text.split('\n')
            enc_camp_text.select()        
        except:
            return "Nothing finded"