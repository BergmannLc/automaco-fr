# -*- coding: utf-8 -*-
import os
import time
import re
import datetime
import pandas as pd
import unicodedata
import pyperclip

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager


# ===== CONFIGURAÇÕES =====
PASTA_BASE = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026"
MANTER_WHATSAPP_ABERTO_NO_FIM = True  # <-- troque pra False se quiser fechar automático

CONTATOS_SETORES = {
    104: ("Jaqueline Castro", "+5524993037830"),
    111: ("Lu", "+5524993032664"),
    125: ("Filipe", "+5524992311056"),
    126: ("Penélope", "+5524992287112"),
    141: ("Jeferson", "+5524992634517"),
    118: ("Jaqueline", "+5524992353744"),
    119: ("José Vicente", "+5524992326541"),
    121: ("Mateus", "+5524992339432"),
    150: ("Caíque Mansur", "+5524992392186"),
    151: ("Álvaro", "+5524992195035"),
    152: ("Cadu", "+5524992011056"),
    153: ("Juliana", "+5524992324258"),
    154: ("Samuel", "+5524993071741"),
    201: ("Daniele", "+5524993032741"),
    202: ("Hamilton Junior", "+5524992466501"),
    114: ("Narjara", "+5524992310680"),
    123: ("Suene", "+5524992313364"),
    133: ("Camilo", "+5524992371891"),
    134: ("Kenia", "+5524993059230"),
    140: ("Marcão", "+5524993003983"),
    301: ("Ian", "+5524992554407"),
    307: ("Guilherme", "+5524992435393"),
    308: ("Jerques", "+5524993049776"),
    401: ("Sebastião", "+5524992484692"),
    402: ("Michel", "+5524992487514"),
    403: ("Robson", "+5524992434226"),
    404: ("Marcos Vinicius", "+5524992482092"),
    405: ("Manoel", "+5524992501104"),
    406: ("Maycon", "+5524993175294"),
    407: ("Alisson", "+5524993068269"),
    501: ("Larissa", "+5524993030179"),
    502: ("Filipe Assis", "+5524993035461"),
    503: ("Janaína", "+5524993065215"),
    504: ("Alexandre", "+5524993053310"),
    505: ("Luís Felipe", "+5524993047559"),
    506: ("Paulo", "+5524992431804"),
    602: ("Joseph", "+5524992316636"),
    603: ("Ana Luiza", "+5524993071548"),
    604: ("Monique", "+5524992346627"),
    605: ("Edson", "+5524992309277"),
    606: ("Victor Pereira", "+5524992430988"),
    113: ("Amanda", "+5524992319906"),
    124: ("Joseni", "+5524992266949"),
    801: ("Victor Prazeres", "+5524993176806"),
    802: ("Lorrane", "+5524992400151"),
    803: ("Aurélio", "+5524992483288"),
    804: ("Davidson", "+5524993060367"),
    805: ("Wellinton", "+5524993057422"),
    807: ("Leandro", "+5524992337480"),
    901: ("Dalso", "+5524993055362"),
    902: ("Rafael Oliveira", "+5524993175185"),
}


# ===== FUNÇÕES GERAIS =====
def normalizar(txt):
    return unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode().strip().lower()

def limpa_ciclo(v):
    try:
        return str(int(float(v)))
    except Exception:
        return re.sub(r"\s+", "", str(v)).strip()

def limpar_numero(numero):
    """Mantém só dígitos para usar na URL do WhatsApp."""
    return re.sub(r"\D", "", str(numero))

def garantir_diretorio_perfil(path):
    pasta = os.path.abspath(os.path.normpath(path))
    os.makedirs(pasta, exist_ok=True)
    return pasta


# ===== SELENIUM / WHATSAPP =====
def preparar_janela(driver):
    try:
        driver.maximize_window()
    except Exception:
        pass

    try:
        driver.set_window_size(1600, 1000)
    except Exception:
        pass

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

def abrir_whatsapp_web(driver, tentativas=3):
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            preparar_janela(driver)
            driver.get("https://web.whatsapp.com/")
            time.sleep(4)
            esperar_whatsapp_pronto(driver, timeout=120)
            return True
        except Exception as e:
            ultimo_erro = e
            print(f"        ⚠ Falha ao abrir WhatsApp ({tentativa}/{tentativas}): {e}")
            time.sleep(2)

    raise TimeoutException(f"Não foi possível abrir corretamente o WhatsApp Web. Último erro: {ultimo_erro}")

def localizar_caixa_mensagem(driver):
    """
    Procura a caixa real onde a mensagem deve ser digitada.
    Evita depender de data-tab fixo.
    """
    seletores = [
        (By.XPATH, '//footer//div[@contenteditable="true" and @role="textbox"]'),
        (By.XPATH, '//footer//div[@contenteditable="true"]'),
        (By.CSS_SELECTOR, "footer div[contenteditable='true'][role='textbox']"),
        (By.CSS_SELECTOR, "footer div[contenteditable='true']"),
        (By.XPATH, '//div[@id="main"]//footer//div[@contenteditable="true"]'),
        (By.XPATH, '//div[contains(@aria-label, "Digite uma mensagem") and @contenteditable="true"]'),
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

def esperar_caixa_mensagem(driver, timeout=40):
    fim = time.time() + timeout
    ultimo_erro = None

    while time.time() < fim:
        try:
            caixa = localizar_caixa_mensagem(driver)
            if caixa:
                return caixa
        except Exception as e:
            ultimo_erro = e
        time.sleep(1)

    raise TimeoutException(f"Caixa de mensagem não encontrada. Último erro: {ultimo_erro}")

def existe_tela_numero_invalido(driver):
    sinais = [
        "O número de telefone compartilhado por url é inválido",
        "Número de telefone compartilhado por URL é inválido",
        "Phone number shared via url is invalid",
        "não está no WhatsApp",
        "isn't on WhatsApp",
    ]
    try:
        texto = driver.page_source.lower()
        return any(s.lower() in texto for s in sinais)
    except Exception:
        return False

def abrir_chat(driver, wait, nome, numero, tentativas=3):
    """
    Abre o chat do WhatsApp diretamente pelo número e confirma
    pela presença real da caixa de mensagem.
    """
    numero_limpo = limpar_numero(numero)
    ultimo_erro = None

    for tentativa in range(1, tentativas + 1):
        try:
            url = f"https://web.whatsapp.com/send?phone={numero_limpo}&app_absent=0"
            driver.get(url)
            time.sleep(3)

            if existe_tela_numero_invalido(driver):
                raise Exception(f"Número inválido ou não encontrado no WhatsApp: {numero}")

            caixa = esperar_caixa_mensagem(driver, timeout=35)

            try:
                clicar_elemento(driver, caixa)
            except Exception:
                pass

            time.sleep(0.8)
            return True

        except Exception as e:
            ultimo_erro = e
            print(f"        ⚠ Tentativa {tentativa}/{tentativas} falhou ao abrir chat de {nome}: {e}")
            time.sleep(2)

    print(f"        ❌ Não consegui confirmar abertura do chat de {nome}. Último erro: {ultimo_erro}")
    return False

def limpar_caixa_mensagem(caixa):
    try:
        caixa.send_keys(Keys.CONTROL, "a")
        time.sleep(0.2)
        caixa.send_keys(Keys.DELETE)
        time.sleep(0.2)
        return True
    except Exception:
        return False

def enviar_mensagem_unica(driver, wait, mensagem):
    """
    Copia o texto completo e envia com CTRL+V.
    Usa localização robusta da caixa de mensagem.
    """
    caixa = esperar_caixa_mensagem(driver, timeout=40)
    clicar_elemento(driver, caixa)
    time.sleep(0.4)

    limpar_caixa_mensagem(caixa)

    pyperclip.copy(mensagem)
    caixa.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.8)
    caixa.send_keys(Keys.ENTER)

    return True


# ===== PLANILHA =====
mes_atual = datetime.datetime.now().strftime("%m")
meses = {
    "01": "JANEIRO", "02": "FEVEREIRO", "03": "MARÇO", "04": "ABRIL", "05": "MAIO", "06": "JUNHO",
    "07": "JULHO", "08": "AGOSTO", "09": "SETEMBRO", "10": "OUTUBRO", "11": "NOVEMBRO", "12": "DEZEMBRO"
}
arquivo_mes = f"{mes_atual} - Fora de Rota automatico - {meses[mes_atual]}.xlsm"
dia_atual = str(datetime.datetime.now().day)

caminho_planilha = os.path.join(PASTA_BASE, arquivo_mes)
if not os.path.exists(caminho_planilha):
    raise FileNotFoundError(f"Planilha do mês não encontrada: {caminho_planilha}")

print(f"📄 Lendo planilha: {caminho_planilha}")
print(f"📌 Aba do dia: {dia_atual}")

df = pd.read_excel(caminho_planilha, sheet_name=dia_atual, header=1)
df.columns = [normalizar(c) for c in df.columns]

colunas_necessarias = {"setor", "sold", "razao social", "dia", "ciclo", "nao autorizado", "retorno"}
faltando = colunas_necessarias - set(df.columns)
if faltando:
    raise ValueError(f"Colunas faltando na aba '{dia_atual}': {sorted(faltando)}")

df = df.rename(columns={
    "setor": "setor",
    "sold": "sold",
    "razao social": "razao social",
    "dia": "dia",
    "ciclo": "ciclo",
    "nao autorizado": "nao autorizado",
    "retorno": "retorno"
})
df = df[["setor", "sold", "razao social", "dia", "ciclo", "nao autorizado", "retorno"]].copy()

df["setor"] = pd.to_numeric(df["setor"], errors="coerce").astype("Int64")
df = df[df["setor"].notna()].copy()

print(f"📊 Linhas lidas: {len(df)}")
print(f"📦 Setores na planilha (únicos): {len(df['setor'].unique())}")


# ===== WHATSAPP =====
perfil = garantir_diretorio_perfil(r"C:\Users\av\Desktop\Automação SAR\ChromeProfile")

options = Options()
options.add_argument(f"user-data-dir={perfil}")
options.add_argument("--start-maximized")
options.add_argument("--disable-notifications")
options.add_argument("--disable-infobars")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

print("\n[1] ⏳ Abrindo WhatsApp Web...")
abrir_whatsapp_web(driver, tentativas=3)
wait = WebDriverWait(driver, 90)
time.sleep(2)
print("        ✅ WhatsApp pronto")

# ===== ENVIO =====
enviados = 0
ignorados_sem_contato = 0
sem_registros = 0
falhas_envio = 0

setores_planilha = sorted([int(x) for x in df["setor"].dropna().unique()])

for setor in setores_planilha:
    if setor not in CONTATOS_SETORES:
        print(f"⚠️ Setor {setor} não está na lista de contatos, ignorando.")
        ignorados_sem_contato += 1
        continue

    nome, numero = CONTATOS_SETORES[setor]
    subset = df[df["setor"] == setor].copy()

    if subset.empty:
        print(f"⚠️ Setor {setor} sem registros.")
        sem_registros += 1
        continue

    header = f"Fora de Rota - Setor {setor} ({datetime.datetime.now():%d/%m})\n\n"
    linhas = [header]

    for i, r in enumerate(subset.itertuples(index=False), 1):
        sold = str(r[1]).strip()
        razao = str(r[2]).strip()
        dia = str(r[3]).strip()
        ciclo = limpa_ciclo(r[4])
        situacao = str(r[5]).strip()
        retorno = r[6]

        bloco = (
            f"{i}) {sold} - {razao}\n"
            f"Dia {dia} | Ciclo {ciclo}\n"
            f"Situação: {situacao}\n"
        )
        if pd.notna(retorno) and str(retorno).strip():
            bloco += f"Retorno: {str(retorno).strip()}\n"
        bloco += "\n"
        linhas.append(bloco)

    mensagem = "".join(linhas).strip()

    print(f"\n📤 Enviando mensagem única para {nome} (Setor {setor})...")

    if abrir_chat(driver, wait, nome, numero):
        try:
            enviar_mensagem_unica(driver, wait, mensagem)
            enviados += 1
            print(f"✅ Mensagem enviada com sucesso para {nome} (Setor {setor})")
        except Exception as e:
            falhas_envio += 1
            print(f"❌ Falha ao enviar para {nome} (Setor {setor}): {e}")
    else:
        falhas_envio += 1
        print(f"❌ Não consegui abrir o chat de {nome} (Setor {setor})")

    time.sleep(2)

print("\n🎉 Envios concluídos!")
print(f"✅ Enviados: {enviados}")
print(f"⚠️ Ignorados (sem contato): {ignorados_sem_contato}")
print(f"⚠️ Sem registros (após filtro): {sem_registros}")
print(f"❌ Falhas de envio: {falhas_envio}")

if MANTER_WHATSAPP_ABERTO_NO_FIM:
    input("\nPressione ENTER para fechar o WhatsApp/Chrome...")

driver.quit()