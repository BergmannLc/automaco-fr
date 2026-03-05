import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
import time
from datetime import date

# =================================================================
#                         CONFIGURAÇÕES
# =================================================================

# Caminho antigo do driver (não é mais usado se você deixar o Selenium Manager cuidar):
DRIVER_PATH = r"C:\Users\AV\Desktop\Automação SAR\chromedriver.exe"

EXCEL_PATH = r"C:\Users\AV\Desktop\Automação SAR\Vendas_Do_Dia.xlsx"
NOME_GRUPO = "Fora de Rota"
USER_DATA_DIR = r"C:\Users\AV\Desktop\Automação SAR\ChromeProfile"

REGEX_CODIGOS = r"\b(\d{5,})\b"
PADRAO_DATA_PRE = re.compile(r'\[[^\]]*(\d{1,2}/\d{1,2}/\d{2,4})')
PADRAO_CONTATO_PRE = re.compile(r']\s*(.*?):\s*$', re.UNICODE)

# =================================================================
#                 FUNÇÕES AUXILIARES
# =================================================================

def extrair_codigo_vendedor(texto):
    if not texto:
        return "Desconhecido"
    texto = texto.replace('\u202a', '').replace('\u202c', '').replace('\u200e', '')
    match = re.match(r'^(\d{3,4})\s*-\s*-?\s*.+', texto)
    if match:
        return match.group(1).strip()
    return "Desconhecido"

def normalizar_nome_contato(texto):
    if not texto:
        return "Desconhecido"
    texto = texto.replace('\u202a', '').replace('\u202c', '').replace('\u200e', '')
    match = re.match(r'^\d{3,4}\s*-\s*-?\s*(.+)', texto)
    if match:
        return match.group(1).strip() or "Desconhecido"
    return texto.strip() or "Desconhecido"

# =================================================================
#                 FUNÇÕES DE CONTROLE DO WHATSAPP
# =================================================================

MAX_TENTATIVAS_SEM_NOVOS = 8
TEMPO_ENTRE_ROLAGENS = 3
SCROLL_DISTANCE = 3

def inferir_data_referencia(driver):
    try:
        main = driver.find_element(By.XPATH, '//div[@id="main"]')
        mensagens = main.find_elements(By.XPATH, './/div[@data-pre-plain-text]')
        for mensagem in reversed(mensagens):
            data_pre = mensagem.get_attribute('data-pre-plain-text') or ''
            match = PADRAO_DATA_PRE.search(data_pre)
            if match:
                return match.group(1)
    except Exception:
        pass
    return date.today().strftime('%d/%m/%Y')

def obter_container_mensagens(driver):
    seletores = [
        (By.CSS_SELECTOR, "#main [data-testid='conversation-panel-body']"),
        (By.XPATH, "//div[@id='main']//div[@role='application']"),
        (By.XPATH, "//div[@id='main']//div[contains(@class, 'copyable-area')]"),
        (By.XPATH, "//div[@id='main']")
    ]
    for by, seletor in seletores:
        try:
            container = driver.find_element(by, seletor)
            if container:
                return container
        except Exception:
            continue
    return driver.find_element(By.TAG_NAME, 'body')

def realizar_rolagem(driver, container):
    try:
        driver.execute_script(
            "arguments[0].scrollTop = arguments[0].scrollTop - (arguments[0].clientHeight * 1.5);",
            container
        )
    except Exception:
        pass
    try:
        ActionChains(driver).move_to_element(container).click(container).send_keys(Keys.PAGE_UP).perform()
    except Exception:
        try:
            container.send_keys(Keys.PAGE_UP)
        except Exception:
            driver.execute_script("window.scrollBy(0, -800);")

# =================================================================
#              EXTRAÇÃO DE CÓDIGOS + CONTATO
# =================================================================

def extrair_codigos_com_contato(driver, data_referencia=None):
    codigos_com_contato = {}
    encontrou_data_antiga = False
    data_alvo = data_referencia or date.today().strftime('%d/%m/%Y')
    
    try:
        main = driver.find_element(By.XPATH, '//div[@id="main"]')
        mensagens = main.find_elements(By.XPATH, './/div[@data-pre-plain-text]')
        
        contato_atual = "Desconhecido"
        ultimo_contato = "Desconhecido"
        
        for mensagem in mensagens:
            data_pre = mensagem.get_attribute('data-pre-plain-text') or ''
            data_match = PADRAO_DATA_PRE.search(data_pre)
            data_msg = data_match.group(1) if data_match else data_alvo
            
            if data_msg != data_alvo:
                encontrou_data_antiga = True
                contato_atual = "Desconhecido"
                ultimo_contato = "Desconhecido"
                continue
            
            contato_match = PADRAO_CONTATO_PRE.search(data_pre)
            if contato_match:
                contato_atual = contato_match.group(1)
                ultimo_contato = contato_atual
            else:
                contato_atual = ultimo_contato
            
            texto_mensagem = mensagem.text or ''
            if not texto_mensagem.strip():
                continue
            
            linhas = [linha.strip() for linha in texto_mensagem.split('\n') if linha.strip()]
            
            for linha in linhas:
                if re.match(r'^\d{3,4}\s*-\s*-?\s*.+', linha):
                    contato_atual = linha
                    ultimo_contato = contato_atual
                    continue
                
                codigos_na_linha = re.findall(REGEX_CODIGOS, linha)
                for codigo in codigos_na_linha:
                    chave = (codigo, contato_atual)
                    if 5 <= len(codigo) <= 10 and chave not in codigos_com_contato:
                        codigos_com_contato[chave] = contato_atual
        
    except Exception as e:
        print(f"[AVISO] Erro na extração com contato: {str(e)[:80]}")
    
    return codigos_com_contato, encontrou_data_antiga

# =================================================================
#                     FUNÇÃO PRINCIPAL
# =================================================================

def iniciar_automacao():
    # Usa Selenium Manager (não passa DRIVER_PATH)
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(service=service, options=options)
    
    try:
        print("=" * 80)
        print("AUTOMAÇÃO SAR - EXTRAÇÃO DE CÓDIGOS DO WHATSAPP")
        print("=" * 80)
        
        driver.get("https://web.whatsapp.com/")
        wait = WebDriverWait(driver, 90)
        
        try:
            search_box = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
            ))
            print("Sessão ativa!")
        except:
            print("Escaneie o QR Code para logar no WhatsApp Web...")
            search_box = wait.until(EC.presence_of_element_located(
                (By.XPATH, '//div[@contenteditable="true"][@data-tab="3"]')
            ))
        
        print(f"\nAbrindo grupo '{NOME_GRUPO}'...")
        search_box.click()
        time.sleep(1)
        search_box.clear()
        search_box.send_keys(NOME_GRUPO)
        time.sleep(3)
        grupo = wait.until(EC.element_to_be_clickable((By.XPATH, f"//span[@title='{NOME_GRUPO}']")))
        grupo.click()
        time.sleep(5)
        
        print("Focando na área de mensagens...")
        area_mensagens = driver.find_element(By.XPATH, '//div[@id="main"]')
        ActionChains(driver).move_to_element(area_mensagens).click().perform()
        time.sleep(2)
        print("Foco na área de mensagens confirmado!")
        
        for _ in range(5):
            try:
                area_msgs = driver.find_element(By.XPATH, '//div[@id="main"]')
                area_msgs.send_keys(Keys.END)
            except:
                pass
            time.sleep(0.5)
        
        print("Posicionado no final da conversa!")
        
        print("\nIniciando extração de códigos...")
        todos_codigos_com_contato = {}
        data_referencia = inferir_data_referencia(driver)
        codigos_iniciais, _ = extrair_codigos_com_contato(driver, data_referencia)
        todos_codigos_com_contato.update(codigos_iniciais)
        
        tentativas_sem_novos = 0
        encontrou_data_antiga = False
        scroll_container = obter_container_mensagens(driver)
        
        while tentativas_sem_novos < MAX_TENTATIVAS_SEM_NOVOS and not encontrou_data_antiga:
            for _ in range(SCROLL_DISTANCE):
                try:
                    realizar_rolagem(driver, scroll_container)
                except StaleElementReferenceException:
                    scroll_container = obter_container_mensagens(driver)
                    realizar_rolagem(driver, scroll_container)
                time.sleep(0.4)
            
            time.sleep(TEMPO_ENTRE_ROLAGENS)
            codigos_atuais, encontrou_antigo = extrair_codigos_com_contato(driver, data_referencia)
            if encontrou_antigo:
                encontrou_data_antiga = True
            
            tamanho_antes = len(todos_codigos_com_contato)
            for chave, contato in codigos_atuais.items():
                if chave not in todos_codigos_com_contato:
                    todos_codigos_com_contato[chave] = contato
            novos = len(todos_codigos_com_contato) - tamanho_antes
            tentativas_sem_novos = 0 if novos > 0 else tentativas_sem_novos + 1
        
        # =================================================================
        # SALVAR PLANILHA COMO NÚMEROS
        # =================================================================
        print("\nSalvando planilha de vendas...")
        if todos_codigos_com_contato:
            dados_filtrados = [(codigo, contato) for (codigo, _), contato in todos_codigos_com_contato.items()]
            df = pd.DataFrame(dados_filtrados, columns=['Código', 'Contato'])
            df['Data'] = date.today().strftime('%d/%m/%Y')
            df['Código_Vendedor'] = df['Contato'].apply(extrair_codigo_vendedor)
            df['Contato'] = df['Contato'].apply(normalizar_nome_contato)
            df = df[['Código', 'Data', 'Contato', 'Código_Vendedor']]

            df['Código'] = pd.to_numeric(df['Código'], errors='coerce')
            df['Código_Vendedor'] = pd.to_numeric(df['Código_Vendedor'], errors='coerce')

            with pd.ExcelWriter(EXCEL_PATH, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Vendas')
                workbook = writer.book
                worksheet = writer.sheets['Vendas']
                formato_num = workbook.add_format({'num_format': '0'})
                worksheet.set_column('A:A', 12, formato_num)
                worksheet.set_column('D:D', 14, formato_num)

            print(f"Planilha salva em: {EXCEL_PATH}")
        else:
            print("Nenhum código encontrado.")
    
    finally:
        print("\nFechando navegador em 10 segundos...")
        time.sleep(10)
        driver.quit()
        print("Navegador fechado.")

# =================================================================
#                         EXECUÇÃO
# =================================================================

if __name__ == "__main__":
    iniciar_automacao()
