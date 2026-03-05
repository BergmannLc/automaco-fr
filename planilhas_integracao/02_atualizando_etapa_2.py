import win32com.client as win32
import shutil
import os
import time

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_EXCEL = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"

# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================
def atualizar_planilha(caminho_arquivo):
    """Abre o Excel, atualiza todas as conexões e salva o arquivo."""

    if not os.path.exists(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    # 🧹 LIMPA CACHE COM (gen_py) - Previne erros de interface COM
    gen_py_path = os.path.join(os.environ['LOCALAPPDATA'], 'Temp', 'gen_py')
    if os.path.exists(gen_py_path):
        try:
            shutil.rmtree(gen_py_path)
            print("🧹 Cache COM (gen_py) limpo.")
        except Exception as e:
            print(f"⚠ Aviso: Não foi possível limpar o cache: {e}")

    print(f"\n🚀 Iniciando: {os.path.basename(caminho_arquivo)}")

    # Inicia o Excel
    excel = win32.DispatchEx('Excel.Application')
    excel.Visible = False  
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False

    try:
        # Abre o arquivo
        wb = excel.Workbooks.Open(caminho_arquivo)

        # ⚡ DESATIVA ATUALIZAÇÃO EM SEGUNDO PLANO (Evita o erro de rejeição)
        print("⚙️ Configurando conexões para modo síncrono...")
        for conn in wb.Connections:
            try:
                # Se for OLEDB (Power Query)
                if conn.Type == 1: 
                    conn.OLEDBConnection.BackgroundQuery = False
                # Se for ODBC
                elif conn.Type == 2:
                    conn.ODBCConnection.BackgroundQuery = False
            except:
                pass 

        # 🔄 Atualiza todas as conexões
        print("🔄 Atualizando consultas (isso pode demorar)...")
        wb.RefreshAll()

        # ⏳ LOOP DE VERIFICAÇÃO (Trata o erro de "Chamada Rejeitada")
        print("⏳ Aguardando conclusão e cálculos...")
        
        max_tentativas = 120 # 120 * 5 segundos = 10 minutos de limite
        tentativa = 0
        
        while tentativa < max_tentativas:
            try:
                # Se o estado for 0, o Excel terminou os cálculos
                if excel.CalculationState == 0:
                    break
            except Exception:
                # Se cair aqui, o Excel está "ocupado" e rejeitou a chamada
                pass
            
            time.sleep(5)
            tentativa += 1

        # Salva e fecha
        wb.Save()
        print("\n✅ Sucesso: Planilha atualizada e salva!")

    except Exception as e:
        print(f"\n❌ Erro durante o processo: {e}")
    
    finally:
        try:
            wb.Close(SaveChanges=True)
        except:
            pass
        excel.Quit()
        # Limpeza de memória
        del wb
        del excel
        print("🔒 Excel encerrado.")

# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    try:
        atualizar_planilha(ARQUIVO_EXCEL)
    except Exception as e:
        print(f"❌ Erro fatal: {e}")