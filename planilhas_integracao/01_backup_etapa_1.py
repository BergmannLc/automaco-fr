import os
import shutil
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES
# ============================================================
ARQUIVO_ORIGINAL = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
PASTA_BACKUP = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\BACKUPS\BACKUPS MARÇO"

# ============================================================
# FUNÇÃO PRINCIPAL
# ============================================================
def criar_backup(arquivo_origem, pasta_destino):
    """Cria um backup do arquivo Excel com data e hora no nome."""

    if not os.path.exists(arquivo_origem):
        raise FileNotFoundError(f"Arquivo original não encontrado: {arquivo_origem}")

    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    # Hora e data atuais
    agora = datetime.now()
    hora_str = agora.strftime("%H.%M")
    data_str = agora.strftime("%d.%m")

    # Nome base do arquivo original
    nome_base = os.path.basename(arquivo_origem)

    # Nome do backup no formato: "16.04 - 28.10 - 10 - Fora de Rota automatico - OUTUBRO.xlsm"
    nome_backup = f"{hora_str} - {data_str} - {nome_base}"

    caminho_backup = os.path.join(pasta_destino, nome_backup)

    # Copia o arquivo
    print(f"Criando backup de:\n{arquivo_origem}")
    print(f"Destino:\n{caminho_backup}\n")

    shutil.copy2(arquivo_origem, caminho_backup)

    print("✅ Backup criado com sucesso!")
    return caminho_backup

# ============================================================
# EXECUÇÃO
# ============================================================
if __name__ == "__main__":
    criar_backup(ARQUIVO_ORIGINAL, PASTA_BACKUP)
