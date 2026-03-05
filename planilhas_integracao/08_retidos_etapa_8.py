import win32com.client as win32
import pythoncom
import os
import time
from datetime import datetime
import openpyxl

# ============================================================
# CONFIGURAÇÕES DINÂMICAS (CORRIGIDAS)
# ============================================================

# Captura data atual
agora = datetime.now()
ano_atual = agora.year
mes_num = agora.strftime("%m")  # Pega "01", "02", etc.
mes_nome_en = agora.strftime("%B").upper()

# Tradução para o padrão das suas pastas
mes_traduzido = {
    "JANUARY": "JANEIRO", "FEBRUARY": "FEVEREIRO", "MARCH": "MARÇO",
    "APRIL": "ABRIL", "MAY": "MAIO", "JUNE": "JUNHO",
    "JULY": "JULHO", "AUGUST": "AGOSTO", "SEPTEMBER": "SETEMBRO",
    "OCTOBER": "OUTUBRO", "NOVEMBER": "NOVEMBRO", "DECEMBER": "DEZEMBRO"
}.get(mes_nome_en, "MÊS_DESCONHECIDO")

# Caminhos dos arquivos
ARQUIVO_RETIDOS = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\RETIDOS DIARIO\retidos.BI.xlsx"

# CORREÇÃO: Usamos {mes_num} para o prefixo (01, 02...) e {mes_traduzido} para o nome
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

    # Ler planilha de retidos
    wb_retidos = openpyxl.load_workbook(caminho_retidos, data_only=True)
    ws_retidos = wb_retidos.active

    codigos_retidos = []
    for row in ws_retidos.iter_rows(min_row=2, values_only=True):
        valor = row[5]  # Coluna F
        if valor not in (None, "", " "):
            codigos_retidos.append(valor)
    wb_retidos.close()

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
                ws_coord.Cells(i, 11).Value = int(codigo) # Coluna K
            except:
                ws_coord.Cells(i, 11).Value = str(codigo)

        wb_destino.Save()
        wb_destino.Close(SaveChanges=True)
        print(f"✅ Processo concluído com sucesso!")
    except Exception as e:
        print(f"❌ Erro: {e}")
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    print(f"=== COLAGEM DE RETIDOS | {mes_traduzido} {ano_atual} ===\n")
    colar_retidos(ARQUIVO_RETIDOS, ARQUIVO_DESTINO, ABA_COORDENADAS)