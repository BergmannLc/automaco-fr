# -*- coding: utf-8 -*-
import os
import time
import re
import datetime
import pandas as pd
import unicodedata
import pyperclip

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options


# ===== CONFIGURAÇÕES =====
PASTA_BASE = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026"
MANTER_WHATSAPP_ABERTO_NO_FIM = True  # <-- troque pra False se quiser fechar automático

CONTATOS_SETORES = {
    104: ("Jaqueline Castro", "+5524993037830"),
    111: ("Lu", "+5524993032664"),
    125: ("Filipe", "+5524992311056"),
    126: ("Hamilton Junior", "+5524992287112"),
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
    202: ("", "+5524992466501"),
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


# ===== FUNÇÕES =====
def normalizar(txt):
    return unicodedata.normalize("NFKD", str(txt)).encode("ASCII", "ignore").decode().strip().lower()

def limpa_ciclo(v):
    try:
        return str(int(float(v)))
    except Exception:
        return re.sub(r"\s+", "", str(v)).strip()

def abrir_chat(driver, wait, nome, numero):
    """
    Abre o chat do WhatsApp via busca pelo NOME.
    Se não achar, abre por número.
    """
    try:
        search = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")
        ))
        search.click()
        search.send_keys(Keys.CONTROL, 'a')
        search.send_keys(Keys.DELETE)

        if nome:
            search.send_keys(nome)
            time.sleep(1.2)
            search.send_keys(Keys.ENTER)
        else:
            driver.get(f"https://web.whatsapp.com/send?phone={numero}")

        wait.until(EC.presence_of_element_located(
            (By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
        ))
        return True
    except Exception:
        # fallback por número
        try:
            driver.get(f"https://web.whatsapp.com/send?phone={numero}")
            wait.until(EC.presence_of_element_located(
                (By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
            ))
            return True
        except Exception:
            return False

def enviar_mensagem_unica(driver, wait, mensagem):
    """Copia o texto completo e envia com CTRL+V (garante envio único)."""
    chat_box = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
    ))
    pyperclip.copy(mensagem)
    chat_box.click()
    time.sleep(0.3)
    chat_box.send_keys(Keys.CONTROL, 'v')
    time.sleep(0.6)
    chat_box.send_keys(Keys.ENTER)


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

# valida colunas
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

# ✅ CORREÇÃO PRINCIPAL: normaliza SETOR para inteiro (evita '118.0' vs '118')
df["setor"] = pd.to_numeric(df["setor"], errors="coerce").astype("Int64")

# opcional: remove linhas sem setor
df = df[df["setor"].notna()].copy()

# diagnóstico rápido
print(f"📊 Linhas lidas: {len(df)}")
print(f"📦 Setores na planilha (únicos): {len(df['setor'].unique())}")


# ===== WHATSAPP =====
options = Options()
options.add_argument(r"user-data-dir=C:\Users\av\Desktop\Automação SAR\ChromeProfile")
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(), options=options)
driver.get("https://web.whatsapp.com")

wait = WebDriverWait(driver, 90)
wait.until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true'][@data-tab='3']")))
time.sleep(1)

# ===== ENVIO =====
enviados = 0
ignorados_sem_contato = 0
sem_registros = 0

setores_planilha = sorted([int(x) for x in df["setor"].dropna().unique()])

for setor in setores_planilha:
    if setor not in CONTATOS_SETORES:
        print(f"⚠️ Setor {setor} não está na lista de contatos, ignorando.")
        ignorados_sem_contato += 1
        continue

    nome, numero = CONTATOS_SETORES[setor]
    subset = df[df["setor"] == setor].copy()  # ✅ filtro correto

    if subset.empty:
        print(f"⚠️ Setor {setor} sem registros.")
        sem_registros += 1
        continue

    header = f"Fora de Rota - Setor {setor} ({datetime.datetime.now():%d/%m})\n\n"
    linhas = [header]

    for i, r in enumerate(subset.itertuples(index=False), 1):
        # colunas: setor, sold, razao, dia, ciclo, situacao, retorno
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

    print(f"📤 Enviando mensagem única para {nome} (Setor {setor})...")

    if abrir_chat(driver, wait, nome, numero):
        try:
            enviar_mensagem_unica(driver, wait, mensagem)
            enviados += 1
            print(f"✅ Mensagem enviada com sucesso para {nome} (Setor {setor})")
        except Exception as e:
            print(f"❌ Falha ao enviar para {nome} (Setor {setor}): {e}")
    else:
        print(f"❌ Não consegui abrir o chat de {nome} (Setor {setor})")

    time.sleep(2)

print("\n🎉 Envios concluídos!")
print(f"✅ Enviados: {enviados}")
print(f"⚠️ Ignorados (sem contato): {ignorados_sem_contato}")
print(f"⚠️ Sem registros (após filtro): {sem_registros}")

if MANTER_WHATSAPP_ABERTO_NO_FIM:
    input("\nPressione ENTER para fechar o WhatsApp/Chrome...")
driver.quit()
