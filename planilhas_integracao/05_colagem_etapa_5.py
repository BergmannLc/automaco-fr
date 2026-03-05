import os
import time
from datetime import datetime

import openpyxl
import pythoncom
import win32com.client as win32


# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_VENDAS = r"C:\Users\av\Desktop\Automação SAR\Vendas_Do_Dia.xlsx"
ARQUIVO_DESTINO = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def abrir_excel_com_retry(tentativas=8, espera_seg=1.0):
    """
    Abre uma instância nova do Excel via COM com retries.
    Evita:
      - erro de makepy do EnsureDispatch
      - 'A chamada foi rejeitada pelo chamado'
    """
    pythoncom.CoInitialize()

    last_err = None
    for n in range(1, tentativas + 1):
        try:
            # DispatchEx cria uma instância NOVA do Excel (melhor que Dispatch)
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
def colar_dados(vendas_path, destino_path):
    """Copia dados da planilha Vendas_Do_Dia.xlsx e cola na aba do dia vigente no arquivo destino."""

    # === Verificações ===
    if not os.path.exists(vendas_path):
        raise FileNotFoundError(f"Arquivo de vendas não encontrado: {vendas_path}")
    if not os.path.exists(destino_path):
        raise FileNotFoundError(f"Arquivo destino não encontrado: {destino_path}")

    # === Identificar dia vigente ===
    dia_atual = datetime.now().day
    nome_aba_dia = str(dia_atual)
    print(f"Iniciando colagem na aba '{nome_aba_dia}'...\n")

    # === Ler os dados da planilha de vendas ===
    wb_vendas = openpyxl.load_workbook(vendas_path, data_only=True)
    ws_vendas = wb_vendas.active

    codigos_clientes = []
    codigos_vendedores = []

    for row in ws_vendas.iter_rows(min_row=2, values_only=True):
        if row[0] and row[3]:  # colunas A e D (Código e Código_Vendedor)
            codigos_clientes.append(row[0])
            codigos_vendedores.append(row[3])

    wb_vendas.close()

    print(f"→ {len(codigos_clientes)} registros encontrados na planilha de vendas.\n")

    # === Abrir Excel de destino (COM robusto) ===
    excel = abrir_excel_com_retry()
    wb_destino = excel.Workbooks.Open(destino_path)

    try:
        # Selecionar aba do dia vigente
        ws_dia = wb_destino.Sheets(nome_aba_dia)
    except Exception:
        print(f"❌ Aba '{nome_aba_dia}' não encontrada. Execute primeiro o script 03_copia_etapa_3.py.")
        wb_destino.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
        return

    # === Colar os dados ===
    print("Colando dados...")

    linha_inicio = 3

    # Colar códigos de vendedor (em A)
    for i, valor in enumerate(codigos_vendedores, start=linha_inicio):
        ws_dia.Cells(i, 1).Value = valor  # Coluna A

    # Colar códigos de cliente (em C)
    for i, valor in enumerate(codigos_clientes, start=linha_inicio):
        ws_dia.Cells(i, 3).Value = valor  # Coluna C

    print("✅ Dados colados com sucesso!")

    # === Salvar e fechar ===
    wb_destino.Save()
    wb_destino.Close(SaveChanges=True)
    excel.Quit()
    pythoncom.CoUninitialize()

    print("\n✅ Planilha atualizada e Excel fechado com sucesso!")


# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    colar_dados(ARQUIVO_VENDAS, ARQUIVO_DESTINO)
