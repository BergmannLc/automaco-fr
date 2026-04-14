import win32com.client as win32
import pythoncom
import os
import time
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES DINÂMICAS
# ============================================================

# Captura data atual
agora = datetime.now()
ano_atual = agora.year
mes_num = agora.strftime("%m")  # "01", "02", etc.
mes_nome_en = agora.strftime("%B").upper()

# Tradução para o padrão das pastas
mes_traduzido = {
    "JANUARY": "JANEIRO",
    "FEBRUARY": "FEVEREIRO",
    "MARCH": "MARÇO",
    "APRIL": "ABRIL",
    "MAY": "MAIO",
    "JUNE": "JUNHO",
    "JULY": "JULHO",
    "AUGUST": "AGOSTO",
    "SEPTEMBER": "SETEMBRO",
    "OCTOBER": "OUTUBRO",
    "NOVEMBER": "NOVEMBRO",
    "DECEMBER": "DEZEMBRO"
}.get(mes_nome_en, "MÊS_DESCONHECIDO")

# Caminhos dos arquivos
ARQUIVO_RETIDOS = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\RETIDOS DIARIO\retidos.txt"
ARQUIVO_DESTINO = fr"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - {ano_atual}\{mes_num} - Fora de Rota automatico - {mes_traduzido}.xlsm"

ABA_COORDENADAS = "COORDENADAS"

# ============================================================
# AUXILIAR: abrir Excel
# ============================================================
def abrir_excel_com_retry(tentativas=8, espera_seg=1.0):
    pythoncom.CoInitialize()
    last_err = None

    for _ in range(tentativas):
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            return excel
        except Exception as e:
            last_err = e
            time.sleep(espera_seg)

    raise last_err

# ============================================================
# AUXILIAR: ler códigos do TXT
# ============================================================
def ler_codigos_txt(caminho_txt):
    codigos = []

    with open(caminho_txt, "r", encoding="utf-8") as arquivo:
        for linha in arquivo:
            valor = linha.strip()

            # Ignora linhas vazias
            if valor in ("", " "):
                continue

            codigos.append(valor)

    return codigos

# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================
def colar_retidos(caminho_retidos, caminho_destino, aba_coordenadas):
    # Verificação de existência
    if not os.path.exists(caminho_retidos):
        print(f"❌ Arquivo de retidos não encontrado: {caminho_retidos}")
        return

    if not os.path.exists(caminho_destino):
        print(f"❌ Arquivo destino não encontrado: {caminho_destino}")
        print(f"💡 DICA: O script tentou buscar o prefixo '{mes_num}'. Verifique o nome na pasta.")
        return

    print(f"📘 Lendo códigos de: {os.path.basename(caminho_retidos)}")

    # Ler códigos do TXT
    try:
        codigos_retidos = ler_codigos_txt(caminho_retidos)
    except Exception as e:
        print(f"❌ Erro ao ler o arquivo TXT: {e}")
        return

    if not codigos_retidos:
        print("⚠ Nenhum código encontrado.")
        return

    print(f"✅ {len(codigos_retidos)} códigos encontrados. Abrindo Excel...")

    # Abrir destino via COM
    excel = abrir_excel_com_retry()

    try:
        wb_destino = excel.Workbooks.Open(caminho_destino)
        ws_coord = wb_destino.Sheets(aba_coordenadas)

        linha_inicial = 8

        for i, codigo in enumerate(codigos_retidos, start=linha_inicial):
            try:
                ws_coord.Cells(i, 11).Value = int(str(codigo).strip())  # Coluna K
            except Exception:
                ws_coord.Cells(i, 11).Value = str(codigo).strip()

        wb_destino.Save()
        wb_destino.Close(SaveChanges=True)
        print("✅ Processo concluído com sucesso!")

    except Exception as e:
        print(f"❌ Erro: {e}")

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    print(f"=== COLAGEM DE RETIDOS | {mes_traduzido} {ano_atual} ===\n")
    colar_retidos(ARQUIVO_RETIDOS, ARQUIVO_DESTINO, ABA_COORDENADAS)