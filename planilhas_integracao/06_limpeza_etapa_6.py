import win32com.client as win32
import pythoncom
import os
import time
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_DESTINO = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"


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
def remover_linhas_vazias(caminho_arquivo):
    """Abre a aba do dia vigente e remove as linhas que não possuem dados retornados (códigos inválidos)."""

    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    # Nome da aba do dia vigente
    dia_atual = datetime.now().day
    nome_aba_dia = str(dia_atual)
    print(f"Iniciando limpeza de códigos inválidos na aba '{nome_aba_dia}'...\n")

    excel = abrir_excel_com_retry()
    wb = excel.Workbooks.Open(caminho_arquivo)

    try:
        ws = wb.Sheets(nome_aba_dia)
    except Exception:
        print(f"❌ Aba '{nome_aba_dia}' não encontrada. Execute primeiro os scripts anteriores.")
        wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
        return

    # Descobre a última linha usada na coluna C (onde estão os códigos)
    # xlUp = -4162
    ultima_linha = ws.Cells(ws.Rows.Count, 3).End(-4162).Row

    print(f"Verificando linhas de 3 até {ultima_linha}...")

    linhas_excluidas = 0

    # Percorre de baixo pra cima (evita erro ao deletar)
    for linha in range(ultima_linha, 2, -1):
        codigo = ws.Cells(linha, 3).Value      # coluna C (código)
        razao_social = ws.Cells(linha, 4).Value  # coluna D (razão social)

        # Se a célula de "Razão Social" estiver vazia, exclui a linha
        if codigo is not None and (razao_social is None or str(razao_social).strip() == ""):
            ws.Rows(linha).Delete()
            linhas_excluidas += 1

    wb.Save()
    wb.Close(SaveChanges=True)
    excel.Quit()
    pythoncom.CoUninitialize()

    print(f"\n✅ Limpeza concluída. {linhas_excluidas} linhas inválidas foram removidas com sucesso!")


# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    remover_linhas_vazias(ARQUIVO_DESTINO)
