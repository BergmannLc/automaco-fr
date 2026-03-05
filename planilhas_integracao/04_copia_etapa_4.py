import win32com.client as win32
import pythoncom
import os
import time
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_EXCEL = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
ABA_MODELO = "MODELO"
ABA_REFERENCIA = "FORA DE ROTA"  # nova aba será criada logo após esta


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
    for n in range(1, tentativas + 1):
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
def criar_copia_dia(caminho_arquivo, aba_modelo, aba_referencia):
    """Copia a aba MODELO, renomeia com o dia atual e posiciona após a aba FORA DE ROTA.
       Caso a aba já exista, ela é sobrescrita automaticamente.
    """

    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    dia_atual = datetime.now().day
    nome_nova_aba = str(dia_atual)

    print(f"Iniciando criação/sobrescrita da aba '{nome_nova_aba}' com base em '{aba_modelo}'...\n")

    excel = abrir_excel_com_retry()
    wb = excel.Workbooks.Open(caminho_arquivo)

    try:
        ws_modelo = wb.Sheets(aba_modelo)

        # Localiza a aba de referência
        try:
            ws_referencia = wb.Sheets(aba_referencia)
        except Exception:
            print(f"⚠️ Aba de referência '{aba_referencia}' não encontrada. A nova aba será criada no final.")
            ws_referencia = None

        # Se já existir uma aba com o nome do dia, apaga antes
        for ws in wb.Sheets:
            if ws.Name == nome_nova_aba:
                print(f"ℹ️ A aba '{nome_nova_aba}' já existe. Excluindo antes de recriar...")
                excel.DisplayAlerts = False
                ws.Delete()
                excel.DisplayAlerts = False
                break

        # Cria cópia do MODELO
        if ws_referencia:
            ws_modelo.Copy(After=ws_referencia)
            nova_aba = wb.Sheets(ws_referencia.Index + 1)
        else:
            ws_modelo.Copy(After=wb.Sheets(wb.Sheets.Count))
            nova_aba = wb.Sheets(wb.Sheets.Count)

        nova_aba.Name = nome_nova_aba
        print(f"✅ Aba '{nome_nova_aba}' criada após '{aba_referencia}' com sucesso!")

        # === COLORIR AS ABAS ===
        vermelho_escuro = 192 + (0 << 8) + (0 << 16)   # RGB(192,0,0)
        cinza_escuro = 64 + (64 << 8) + (64 << 16)     # RGB(64,64,64)

        for ws in wb.Sheets:
            if ws.Name.isdigit():
                if ws.Name == nome_nova_aba:
                    ws.Tab.Color = vermelho_escuro
                else:
                    ws.Tab.Color = cinza_escuro

        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()

        print("\n✅ Aba criada, colorida e posicionada corretamente!")

    except Exception as e:
        print(f"❌ Erro ao criar ou posicionar a aba: {e}")
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
    criar_copia_dia(ARQUIVO_EXCEL, ABA_MODELO, ABA_REFERENCIA)
