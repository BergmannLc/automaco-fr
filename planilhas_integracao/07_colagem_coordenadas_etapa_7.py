import win32com.client as win32
import pythoncom
import os
import time
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_EXCEL = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
ABA_COORDENADAS = "COORDENADAS"


# ============================================================
# AUXILIAR: abrir Excel sem EnsureDispatch + retry
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

    for _ in range(tentativas):
        try:
            excel = win32.DispatchEx("Excel.Application")  # instância nova
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
def copiar_codigos_para_coordenadas(caminho_arquivo, aba_coordenadas):
    """Copia os códigos visíveis da aba do dia vigente (coluna C) e cola na aba COORDENADAS (coluna S, linha 8)."""

    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    # Dia vigente
    dia_atual = datetime.now().day
    nome_aba_dia = str(dia_atual)

    print(f"Iniciando cópia de códigos da aba '{nome_aba_dia}' para '{aba_coordenadas}'...\n")

    excel = abrir_excel_com_retry()
    wb = excel.Workbooks.Open(caminho_arquivo)

    try:
        ws_dia = wb.Sheets(nome_aba_dia)
        ws_coord = wb.Sheets(aba_coordenadas)

        # Força cálculo de todas as fórmulas antes de copiar
        print("Recalculando fórmulas antes de copiar...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()
        wb.Application.CalculateFull()

        # Última linha usada na coluna C da aba do dia
        ultima_linha = ws_dia.Cells(ws_dia.Rows.Count, 3).End(-4162).Row  # xlUp = -4162
        qtd = ultima_linha - 2  # começa na linha 3

        if qtd <= 0:
            print("❌ Nenhum código encontrado na aba do dia vigente.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()
            return

        print(f"Copiando {qtd} códigos da aba '{nome_aba_dia}'...")

        # Lê os valores da coluna C
        codigos = []
        for i in range(3, ultima_linha + 1):
            valor = ws_dia.Cells(i, 3).Value
            if valor not in (None, "", " "):
                codigos.append(valor)

        if not codigos:
            print("❌ Nenhum valor válido encontrado para copiar.")
            wb.Close(SaveChanges=False)
            excel.Quit()
            pythoncom.CoUninitialize()
            return

        # Cola na aba COORDENADAS, coluna S a partir da linha 8
        linha_destino = 8
        for i, codigo in enumerate(codigos, start=linha_destino):
            ws_coord.Cells(i, 19).Value = codigo  # Coluna S = 19

        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()

        print(f"✅ {len(codigos)} códigos copiados da aba '{nome_aba_dia}' para '{aba_coordenadas}' com sucesso!")

    except Exception as e:
        print(f"❌ Erro durante o processo: {e}")
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
        excel.Quit()
        pythoncom.CoUninitialize()


# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    copiar_codigos_para_coordenadas(ARQUIVO_EXCEL, ABA_COORDENADAS)
