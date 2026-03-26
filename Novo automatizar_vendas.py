import os
import re
import time
from datetime import date

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import (
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# =================================================================
#                         CONFIGURAÇÕES
# =================================================================

EXCEL_PATH = r"C:\Users\AV\Desktop\Automação SAR\Vendas_Do_Dia.xlsx"
NOME_GRUPO = "Fora de Rota"
USER_DATA_DIR = r"C:\Users\AV\Desktop\Automação SAR\ChromeProfile"

REGEX_CODIGOS = r"\b(\d{5,})\b"
PADRAO_DATA_PRE = re.compile(r'\[[^\]]*(\d{1,2}/\d{1,2}/\d{2,4})')
PADRAO_CONTATO_PRE = re.compile(r']\s*(.*?):\s*$', re.UNICODE)

MAX_TENTATIVAS_SEM_NOVOS = 8
TEMPO_ENTRE_ROLAGENS = 3
SCROLL_DISTANCE = 3

# =================================================================
#                         HELPERS GERAIS
# =================================================================

def garantir_diretorio_perfil(path):
    pasta = os.path.abspath(os.path.normpath(path))
    os.makedirs(pasta, exist_ok=True)
    return pasta

def limpar_texto_oculto(texto):
    if not texto:
        return ""
    return (
        texto.replace("\u202a", "")
        .replace("\u202c", "")
        .replace("\u200e", "")
        .strip()
    )

def extrair_codigo_vendedor(texto):
    texto = limpar_texto_oculto(texto)
    if not texto:
        return "Desconhecido"

    match = re.match(r'^(\d{3,4})\s*-\s*-?\s*.+', texto)
    if match:
        return match.group(1).strip()

    return "Desconhecido"

def normalizar_nome_contato(texto):
    texto = limpar_texto_oculto(texto)
    if not texto:
        return "Desconhecido"

    match = re.match(r'^\d{3,4}\s*-\s*-?\s*(.+)', texto)
    if match:
        return match.group(1).strip() or "Desconhecido"

    return texto.strip() or "Desconhecido"

def _data_hoje_strings():
    hoje = date.today()
    return hoje.strftime("%d/%m/%Y"), hoje.strftime("%d/%m/%y")

def _eh_mesma_data(data_msg, data_alvo):
    if not data_msg or not data_alvo:
        return False

    formatos = ["%d/%m/%Y", "%d/%m/%y"]

    data_msg_dt = None
    data_alvo_dt = None

    for fmt in formatos:
        try:
            data_msg_dt = data_msg_dt or __import__("datetime").datetime.strptime(data_msg, fmt).date()
        except Exception:
            pass
        try:
            data_alvo_dt = data_alvo_dt or __import__("datetime").datetime.strptime(data_alvo, fmt).date()
        except Exception:
            pass

    if data_msg_dt and data_alvo_dt:
        return data_msg_dt == data_alvo_dt

    return data_msg == data_alvo

# =================================================================
#                 FUNÇÕES DE CONTROLE DO WHATSAPP
# =================================================================

def preparar_janela(driver):
    try:
        driver.maximize_window()
    except Exception:
        pass

    try:
        driver.set_window_size(1600, 1000)
    except Exception:
        pass

def esperar_whatsapp_pronto(driver, timeout=120):
    wait = WebDriverWait(driver, timeout)
    candidatos = [
        (By.ID, "side"),
        (By.ID, "pane-side"),
        (By.XPATH, '//*[contains(text(), "Pesquisar ou começar uma nova conversa")]'),
        (By.XPATH, '//*[contains(@aria-label, "Pesquisar")]'),
        (By.XPATH, '//*[contains(@title, "Pesquisar")]'),
        (By.XPATH, '//div[@id="main"]'),
    ]

    ultimo_erro = None
    for by, sel in candidatos:
        try:
            wait.until(EC.presence_of_element_located((by, sel)))
            return True
        except Exception as e:
            ultimo_erro = e
            continue

    raise TimeoutException(f"WhatsApp Web não ficou pronto. Último erro: {ultimo_erro}")

def clicar_elemento(driver, elemento):
    try:
        elemento.click()
        return True
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(elemento).pause(0.2).click(elemento).perform()
        return True
    except Exception:
        pass

    try:
        driver.execute_script("arguments[0].click();", elemento)
        return True
    except Exception:
        pass

    return False

def abrir_whatsapp_web(driver, tentativas=3):
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            preparar_janela(driver)

            try:
                driver.switch_to.new_window("tab")
            except Exception:
                driver.execute_script("window.open('about:blank','_blank');")
                driver.switch_to.window(driver.window_handles[-1])

            preparar_janela(driver)
            driver.get("https://web.whatsapp.com/")
            time.sleep(4)

            preparar_janela(driver)
            esperar_whatsapp_pronto(driver, timeout=120)
            return True

        except Exception as e:
            ultimo_erro = e
            try:
                if len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[-1])
            except Exception:
                pass
            time.sleep(2)

    raise TimeoutException(f"Não foi possível abrir corretamente o WhatsApp Web. Último erro: {ultimo_erro}")

def localizar_barra_busca_visual(driver):
    seletores = [
        (By.XPATH, '//*[contains(text(), "Pesquisar ou começar uma nova conversa")]'),
        (By.XPATH, '//*[contains(@title, "Pesquisar ou começar uma nova conversa")]'),
        (By.XPATH, '//*[contains(@aria-label, "Pesquisar ou começar uma nova conversa")]'),
        (By.XPATH, '//*[contains(text(), "Pesquisar")]'),
        (By.XPATH, '//*[contains(@aria-label, "Pesquisar")]'),
        (By.XPATH, '//*[contains(@title, "Pesquisar")]'),
        (By.XPATH, '//div[@id="side"]//div[@role="button"]'),
    ]

    for by, sel in seletores:
        try:
            elems = driver.find_elements(by, sel)
            for el in elems:
                if el.is_displayed():
                    return el
        except Exception:
            continue
    return None

def localizar_input_busca(driver):
    seletores = [
        (By.XPATH, '//div[@id="side"]//div[@contenteditable="true" and @role="textbox"]'),
        (By.XPATH, '//div[@id="side"]//div[@contenteditable="true"]'),
        (By.CSS_SELECTOR, 'div#side div[contenteditable="true"]'),
        (By.XPATH, '//div[@contenteditable="true" and @role="textbox"]'),
        (By.XPATH, '//div[@contenteditable="true"]'),
    ]

    for by, sel in seletores:
        try:
            elems = driver.find_elements(by, sel)
            for el in elems:
                if el.is_displayed() and el.is_enabled():
                    return el
        except Exception:
            continue
    return None

def limpar_estado_busca(driver):
    try:
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.3)
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.3)
    except Exception:
        pass

def set_text_contenteditable_via_js(driver, elemento, texto):
    script = """
    const el = arguments[0];
    const text = arguments[1];

    el.focus();

    if (el.getAttribute('contenteditable') === 'true') {
        el.innerHTML = '';
        el.textContent = text;

        el.dispatchEvent(new InputEvent('input', {
            data: text,
            inputType: 'insertText',
            bubbles: true,
            cancelable: true
        }));

        el.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
    }

    if (el.tagName === 'INPUT' || el.tagName === 'TEXTAREA') {
        el.value = text;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
    }

    return false;
    """
    try:
        return driver.execute_script(script, elemento, texto)
    except Exception:
        return False

def ativar_e_digitar_na_pesquisa(driver, texto, timeout=30):
    fim = time.time() + timeout
    ultimo_erro = None

    while time.time() < fim:
        try:
            limpar_estado_busca(driver)

            barra_visual = localizar_barra_busca_visual(driver)
            if barra_visual:
                clicar_elemento(driver, barra_visual)
                time.sleep(0.8)

            input_real = localizar_input_busca(driver)

            if not input_real:
                try:
                    ativo = driver.switch_to.active_element
                    if ativo and ativo.is_displayed():
                        input_real = ativo
                except Exception:
                    pass

            if not input_real and barra_visual:
                try:
                    clicar_elemento(driver, barra_visual)
                    time.sleep(0.3)
                    barra_visual.send_keys(Keys.TAB)
                    time.sleep(0.5)
                    ativo = driver.switch_to.active_element
                    if ativo and ativo.is_displayed():
                        input_real = ativo
                except Exception:
                    pass

            if not input_real:
                time.sleep(1)
                continue

            try:
                input_real.send_keys(Keys.CONTROL, 'a')
                time.sleep(0.2)
                input_real.send_keys(Keys.DELETE)
                time.sleep(0.2)
                input_real.send_keys(texto)
                time.sleep(2)
                return True
            except Exception:
                pass

            if set_text_contenteditable_via_js(driver, input_real, texto):
                time.sleep(2)
                return True

            try:
                ativo = driver.switch_to.active_element
                if ativo and set_text_contenteditable_via_js(driver, ativo, texto):
                    time.sleep(2)
                    return True
            except Exception:
                pass

        except Exception as e:
            ultimo_erro = e
            time.sleep(1)

    raise TimeoutException(f"Não foi possível ativar/digitar na caixa de pesquisa. Último erro: {ultimo_erro}")

def abrir_chat_por_nome(nome, driver, timeout=45):
    fim = time.time() + timeout
    primeiro_termo = nome.split()[0]

    seletores = [
        (By.XPATH, f'//span[@title="{nome}"]'),
        (By.XPATH, f'//span[contains(@title, "{nome}")]'),
        (By.XPATH, f'//span[contains(@title, "{primeiro_termo}")]'),
        (By.XPATH, f'//div[@id="pane-side"]//span[@title="{nome}"]'),
        (By.XPATH, f'//div[@id="pane-side"]//span[contains(@title, "{primeiro_termo}")]'),
        (By.XPATH, f'//div[@id="side"]//span[@title="{nome}"]'),
    ]

    while time.time() < fim:
        for by, sel in seletores:
            try:
                elems = driver.find_elements(by, sel)
                for el in elems:
                    if el.is_displayed():
                        if clicar_elemento(driver, el):
                            time.sleep(1)
                            return True
            except Exception:
                pass
        time.sleep(1)

    raise TimeoutException(f"Grupo '{nome}' não encontrado")

def buscar_e_abrir_grupo(driver, nome_grupo):
    for tentativa in range(1, 3):
        try:
            print(f"        🔄 Tentativa de busca {tentativa}/2...")
            ativar_e_digitar_na_pesquisa(driver, nome_grupo, timeout=30)
            abrir_chat_por_nome(nome_grupo, driver, timeout=20)
            return True
        except Exception as e:
            print(f"        ⚠ Falha na busca na tentativa {tentativa}: {e}")
            if tentativa < 2:
                print("        🔃 Recarregando WhatsApp Web e tentando novamente...")
                driver.refresh()
                esperar_whatsapp_pronto(driver, timeout=120)
                time.sleep(3)
            else:
                raise

def obter_container_mensagens(driver):
    seletores = [
        (By.CSS_SELECTOR, "#main [data-testid='conversation-panel-body']"),
        (By.XPATH, "//div[@id='main']//div[@role='application']"),
        (By.XPATH, "//div[@id='main']//div[contains(@class, 'copyable-area')]"),
        (By.XPATH, "//div[@id='main']"),
    ]
    for by, sel in seletores:
        try:
            container = driver.find_element(by, sel)
            if container:
                return container
        except Exception:
            continue
    return driver.find_element(By.TAG_NAME, "body")

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

def inferir_data_referencia(driver):
    try:
        main = driver.find_element(By.XPATH, '//div[@id="main"]')
        mensagens = main.find_elements(By.XPATH, './/div[@data-pre-plain-text]')
        for mensagem in reversed(mensagens):
            data_pre = limpar_texto_oculto(mensagem.get_attribute('data-pre-plain-text') or '')
            match = PADRAO_DATA_PRE.search(data_pre)
            if match:
                return match.group(1)
    except Exception:
        pass
    return date.today().strftime('%d/%m/%Y')

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
            data_pre = limpar_texto_oculto(mensagem.get_attribute('data-pre-plain-text') or '')
            data_match = PADRAO_DATA_PRE.search(data_pre)
            data_msg = data_match.group(1) if data_match else data_alvo

            if not _eh_mesma_data(data_msg, data_alvo):
                encontrou_data_antiga = True
                contato_atual = "Desconhecido"
                ultimo_contato = "Desconhecido"
                continue

            contato_match = PADRAO_CONTATO_PRE.search(data_pre)
            if contato_match:
                contato_atual = limpar_texto_oculto(contato_match.group(1))
                ultimo_contato = contato_atual
            else:
                contato_atual = ultimo_contato

            texto_mensagem = limpar_texto_oculto(mensagem.text or '')
            if not texto_mensagem:
                continue

            linhas = [limpar_texto_oculto(linha) for linha in texto_mensagem.split('\n') if limpar_texto_oculto(linha)]

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
        print(f"[AVISO] Erro na extração com contato: {str(e)[:150]}")

    return codigos_com_contato, encontrou_data_antiga

def salvar_debug(driver):
    try:
        main = driver.find_element(By.XPATH, '//div[@id="main"]')
        with open("debug_vendas_fora_rota.txt", "w", encoding="utf-8") as f:
            f.write(main.text)
        return True
    except Exception:
        return False

# =================================================================
#                         EXCEL
# =================================================================

def salvar_planilha_vendas(todos_codigos_com_contato):
    dados_filtrados = [
        (codigo, contato)
        for (codigo, _), contato in todos_codigos_com_contato.items()
    ]

    df = pd.DataFrame(dados_filtrados, columns=['Código', 'Contato'])
    df['Data'] = date.today().strftime('%d/%m/%Y')
    df['Código_Vendedor'] = df['Contato'].apply(extrair_codigo_vendedor)
    df['Contato'] = df['Contato'].apply(normalizar_nome_contato)
    df = df[['Código', 'Data', 'Contato', 'Código_Vendedor']]

    df = df.drop_duplicates(subset=['Código', 'Contato', 'Código_Vendedor']).sort_values(
        by=['Código', 'Contato'],
        key=lambda s: s.astype(str)
    )

    df['Código'] = pd.to_numeric(df['Código'], errors='coerce')
    df['Código_Vendedor'] = pd.to_numeric(df['Código_Vendedor'], errors='coerce')

    with pd.ExcelWriter(EXCEL_PATH, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Vendas')
        workbook = writer.book
        worksheet = writer.sheets['Vendas']

        formato_num = workbook.add_format({'num_format': '0'})
        worksheet.set_column('A:A', 12, formato_num)
        worksheet.set_column('B:B', 12)
        worksheet.set_column('C:C', 35)
        worksheet.set_column('D:D', 14, formato_num)

    return len(df)

# =================================================================
#                     FUNÇÃO PRINCIPAL
# =================================================================

def iniciar_automacao():
    perfil = garantir_diretorio_perfil(USER_DATA_DIR)

    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={perfil}")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    try:
        print("=" * 80)
        print("AUTOMAÇÃO SAR - EXTRAÇÃO DE CÓDIGOS DO WHATSAPP")
        print("=" * 80)

        print("\n[1] ⏳ Abrindo WhatsApp Web...")
        abrir_whatsapp_web(driver, tentativas=3)
        time.sleep(3)
        print("        ✅ Interface pronta")

        print(f"\n[2] 🔍 Buscando grupo '{NOME_GRUPO}'...")
        buscar_e_abrir_grupo(driver, NOME_GRUPO)
        print("        ✅ Grupo aberto!")

        print("\n[3] 📜 Posicionando no fim da conversa...")
        try:
            area = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//div[@id="main"]'))
            )
            ActionChains(driver).move_to_element(area).click(area).perform()
            time.sleep(0.5)
            area.send_keys(Keys.END)
            time.sleep(1)

            for _ in range(4):
                try:
                    area.send_keys(Keys.END)
                except Exception:
                    pass
                time.sleep(0.5)
        except Exception:
            pass
        print("        ✅ Posicionado")

        print("\n[4] 🗓 Definindo data de referência...")
        data_referencia = inferir_data_referencia(driver)
        print(f"        ✅ Data detectada: {data_referencia}")

        print("\n[5] 🔢 Extraindo códigos...")
        todos_codigos_com_contato = {}
        codigos_iniciais, _ = extrair_codigos_com_contato(driver, data_referencia)
        todos_codigos_com_contato.update(codigos_iniciais)
        print(f"        → {len(todos_codigos_com_contato)} códigos iniciais")

        print("\n[6] 🔄 Rolando para trás e capturando mais mensagens...")
        tentativas_sem_novos = 0
        encontrou_data_antiga = False
        scroll_container = obter_container_mensagens(driver)
        contador_rolagens = 0

        while tentativas_sem_novos < MAX_TENTATIVAS_SEM_NOVOS and not encontrou_data_antiga:
            contador_rolagens += 1

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

            if novos > 0:
                tentativas_sem_novos = 0
                print(f"        ✅ Rolagem {contador_rolagens:2d}: +{novos} | Total {len(todos_codigos_com_contato)}")
            else:
                tentativas_sem_novos += 1
                print(f"        ⚪ Rolagem {contador_rolagens:2d}: sem novos ({tentativas_sem_novos}/{MAX_TENTATIVAS_SEM_NOVOS})")

        if encontrou_data_antiga:
            print("        ℹ Mensagens de outra data encontradas. Encerrando coleta.")
        else:
            print("        ℹ Limite de tentativas sem novos atingido.")

        print("\n[7] 💾 Salvando debug...")
        if salvar_debug(driver):
            print("        ✅ Debug salvo")
        else:
            print("        ⚠ Não foi possível salvar debug")

        print("\n[8] 📊 Gerando planilha...")
        if todos_codigos_com_contato:
            qtd = salvar_planilha_vendas(todos_codigos_com_contato)
            print(f"        ✅ Planilha salva com {qtd} linhas em: {EXCEL_PATH}")
        else:
            print("        ❌ Nenhum código encontrado.")

    except Exception as e:
        print("\n" + "=" * 80)
        print("❌ ERRO FATAL NA AUTOMAÇÃO")
        print("=" * 80)
        print(f"\nErro: {str(e)}\n")
        import traceback
        print(traceback.format_exc())

    finally:
        print("\n⏳ Fechando navegador em 10s...")
        time.sleep(10)
        driver.quit()
        print("✅ Fechado com sucesso.")

# =================================================================
#                         EXECUÇÃO
# =================================================================

if __name__ == "__main__":
    iniciar_automacao()