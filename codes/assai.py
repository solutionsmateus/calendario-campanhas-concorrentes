import os
import time
import re
import pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
# Importe 'openpyxl' se for usar Pandas para XLSX
# import openpyxl 

# --- CONFIGURAÇÕES DE DIRETÓRIO ---
# O diretório de saída é lido da variável de ambiente (GitHub Actions)
OUTPUT_DIR = os.environ.get("OUTPUT_DIR", str(Path.home() / "Desktop/Encartes-Concorrentes/Assai"))
ENCARTE_DIR = Path(OUTPUT_DIR)
Path(ENCARTE_DIR).mkdir(parents=True, exist_ok=True) # Garante que o diretório existe

# O nome final do arquivo Excel será o mesmo para todos os scripts rodando na matriz
XLSX_FILE_PATH = ENCARTE_DIR / "campanhas_concorrentes.xlsx" 

# --- Mapeamento de Lojas ---
LOJAS_ESTADOS = {
    "Maranhão": "Assaí Angelim",
    "Bahia": "Assaí Vitória da Conquista",
    # ... (restante das lojas)
}
REGIAO_POR_ESTADO = {
    "Bahia": "Interior", 
}
BASE_URL = "https://www.assai.com.br/ofertas"


# === HEADLESS CHROME ===
def build_headless_chrome():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # ... (outros argumentos)
    user_agent = os.environ.get("HTTP_UA", "Mozilla/5.0 (...)") # Usa a ENV do YAML
    options.add_argument(f"--user-agent={user_agent}")
    options.add_argument("--lang=pt-BR,pt")
    # Usa o 'executable_path' se for necessário, ou deixe o Selenium encontrar
    return webdriver.Chrome(options=options)

driver = build_headless_chrome()
wait = WebDriverWait(driver, 30)

# --- FUNÇÕES DE UTILIDADE ---

def encontrar_data():
    """Encontra e retorna o texto da validade."""
    try:
        data_el = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ofertas-tab-validade")]'))
        )
        texto = data_el.text.strip()
        if texto:
            nome_pasta = re.sub(r'[\\/*?:"<>|\s]', '_', texto)
            return nome_pasta, texto
    except Exception as e:
        print(f"Erro ao encontrar data: {e}")
    return "sem_data", "Data não encontrada"

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

def select_by_visible_text_contains(select_el, target_text, timeout=10):
    # Função auxiliar para selecionar texto que contém a string alvo
    WebDriverWait(driver, timeout).until(lambda d: len(select_el.find_elements(By.TAG_NAME, "option")) > 0)
    sel = Select(select_el)
    opts = select_el.find_elements(By.TAG_NAME, "option")
    alvo_norm = target_text.strip().lower()
    for o in opts:
        if alvo_norm in o.text.strip().lower():
            sel.select_by_visible_text(o.text)
            return True
    return False

# --- FUNÇÕES DE EXTRAÇÃO E SALVAMENTO ---

def save_as_xlsx(data_dict, file_path):
    """Salva um dicionário de dados em um arquivo XLSX, anexando se ele existir."""
    try:
        df_novo = pd.DataFrame([data_dict])
        
        if file_path.exists():
            # Tenta ler o arquivo existente
            df_existente = pd.read_excel(file_path, engine='openpyxl')
            # Anexa os novos dados
            df_final = pd.concat([df_existente, df_novo], ignore_index=True)
        else:
            df_final = df_novo

        df_final.to_excel(file_path, index=False, engine='openpyxl')
        print(f"Dados anexados/salvos em: {file_path}")
    except Exception as e:
        print(f"ERRO ao salvar no Excel: {e}")
        # É crucial ter 'openpyxl' no requirements.txt

def processar_campanhas(jornal_num, estado, loja):
    """Encontra o container de ofertas, extrai o HTML e salva."""
    
    # 1. Encontrar a data de validade
    data_nome, data_texto = encontrar_data()
    
    try:
        # 2. Encontrar o container de ofertas/produtos
        ofertas_container = aguardar_elemento("div.ofertas-slider", timeout=10)
        
        # 3. Extrair o HTML completo do container para que você possa processar depois
        html_ofertas = ofertas_container.get_attribute('outerHTML')
        
        # 4. Preparar os dados para salvar
        dados = {
            'Empresa': 'Assaí',
            'Estado': estado,
            'Loja': loja,
            'Jornal Número': jornal_num,
            'Validade Nome': data_nome,
            'Validade Texto': data_texto,
            'HTML Container': html_ofertas,
            'Data Coleta': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        
        # 5. Salvar/anexar no arquivo Excel
        save_as_xlsx(dados, XLSX_FILE_PATH)
        
    except Exception as e:
        print(f"Não foi possível processar a campanha {jornal_num} para {loja}. Erro: {e}")
        
    # Lógica de iteração (se houver mais páginas de ofertas no carrossel)
    # Mantive a lógica original, mas ela não parece ser de navegação por "página"
    # e sim por "próxima oferta"
    try:
        next_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.slick-next")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", next_button)
        time.sleep(0.5)
        #next_button.click() # NÃO CLIQUE, se o objetivo é extrair o container atual
        # time.sleep(2)
    except:
        pass # Fim do carrossel de ofertas


# --- INÍCIO DO FLUXO PRINCIPAL ---
try:
    print(f"Iniciando raspagem. Output Dir: {ENCARTE_DIR}")
    driver.get(BASE_URL)
    time.sleep(2)

    try:
        clicar_elemento("button.ot-close-icon")
    except:
        pass

    clicar_elemento("a.seletor-loja")
    time.sleep(1)

    for estado, loja in LOJAS_ESTADOS.items():
        print(f"\n--- Processando: {estado} - {loja} ---")

        # Seleção de Estado
        estado_select = aguardar_elemento("select.estado")
        Select(estado_select).select_by_visible_text(estado)
        time.sleep(1)
        
        # Seleção de Região (se aplicável)
        if estado in REGIAO_POR_ESTADO:
            try:
                regiao_select_element = aguardar_elemento("select.regiao", timeout=15)
                Select(regiao_select_element).select_by_visible_text(REGIAO_POR_ESTADO[estado])
                aguardar_elemento("select.loja option[value]", timeout=20)
                time.sleep(0.5)
            except Exception as e:
                print(f"Não foi possível selecionar a região para {estado}: {e}")
                
        # Seleção de Loja
        loja_select = aguardar_elemento("select.loja", timeout=20)
        try:
            Select(loja_select).select_by_visible_text(loja)
        except:
            ok = select_by_visible_text_contains(loja_select, loja)
            if not ok:
                raise RuntimeError(f"Não encontrei a loja '{loja}' no estado {estado}")

        time.sleep(0.8)

        # Confirmação e espera do carregamento das ofertas
        clicar_elemento("button.confirmar")
        time.sleep(1)
        aguardar_elemento("div.ofertas-slider", timeout=30)
        scroll_down_and_up()
        
        # Processa o Jornal 1
        processar_campanhas(1, estado, loja) # Passando estado e loja

        # Tenta processar jornais adicionais (2 e 3)
        for i in range(2, 4):
            try:
                clicar_elemento(f"//button[contains(., 'Jornal de Ofertas {i}')]", By.XPATH)
                time.sleep(3)
                aguardar_elemento("div.ofertas-slider", timeout=30)
                scroll_down_and_up()
                processar_campanhas(i, estado, loja) # Passando estado e loja
            except Exception as e:
                print(f"Jornal {i} indisponível para {loja}: {str(e)}")

        # Volta para o seletor de loja para o próximo loop
        clicar_elemento("a.seletor-loja")
        time.sleep(2)

    print("\nTodos os encartes foram processados e os dados salvos no Excel!")

except Exception as e:
    print(f"\nErro crítico: {str(e)}")
    # Lógica para salvar screenshot em caso de erro
    try:
        driver.save_screenshot(str(ENCARTE_DIR / f"erro_assai_{datetime.now().timestamp()}.png"))
    except:
        pass
finally:
    driver.quit()
