import win32com.client as win32
import os
import time

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_EXCEL = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
MACRO_NOME = "LIMPAR_DADOS"  # nome da macro VBA
ABA_MACRO = "COORDENADAS"   # aba onde a macro deve rodar

# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================
def executar_macro(caminho_arquivo, nome_macro, aba_macro):
    """Abre o Excel, ativa a aba correta, executa a macro e salva."""

    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    print(f"Iniciando execução da macro '{nome_macro}' na planilha:")
    print(caminho_arquivo + "\n")

    # Inicia o Excel
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # coloque True se quiser ver o Excel abrindo

    # Abre o arquivo
    wb = excel.Workbooks.Open(caminho_arquivo)

    try:
        # Ativa a aba COORDENADAS antes de rodar a macro
        print(f"Ativando aba '{aba_macro}'...")
        ws = wb.Sheets(aba_macro)
        ws.Activate()

        # Executa a macro
        print(f"Executando macro: {nome_macro} ...")
        excel.Application.Run(f"'{os.path.basename(caminho_arquivo)}'!{nome_macro}")
        print("✅ Macro executada com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao executar a macro: {e}")
    finally:
        # Dá um tempinho para finalizar cálculos e salvar
        time.sleep(2)
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()

    print("\n✅ Execução concluída e planilha salva com sucesso!")

# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    executar_macro(ARQUIVO_EXCEL, MACRO_NOME, ABA_MACRO)
