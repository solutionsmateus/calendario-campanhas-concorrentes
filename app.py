import pathlib
import os
import socket
from selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome import remote_connection

#Initial Configuration to run in paralel
ENCARTE_DIR = "C:\Users\pedro.gomes\Desktop\Campanhas_get\Campanhas_doc.docx"

BASE_URL = {
    "Assaí": "https://www.assai.com.br/ofertas",
    "Atacadão": "https://www.atacadao.com.br/institucional/nossas-lojas",
    "Cometa_Supermercados": "https://cometasupermercados.com.br/ofertas/",
    "Atakarejo": "https://atakarejo.com.br/cidade/vitoria-da-conquista",
    "Frangolandia": "https://frangolandia.com/encartes/",
    "GBarbosa": "https://blog.gbarbosa.com.br/ofertas/",
    "Novo_Atacarejo": "https://novoatacarejo.com/oferta/"
}

webdriver.Chrome(BASE_URL)
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)


#Atacadão

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

    driver.get("https://www.atacadao.com.br/institucional/nossas-lojas")
    
    try:
        enc_