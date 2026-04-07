import os
import shutil
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES DINÂMICAS (NOVO PADRÃO)
# ============================================================

# Captura data atual
agora = datetime.now()
ano_atual = agora.year
mes_num = agora.strftime("%m")
mes_nome_en = agora.strftime("%B").upper()

# Tradução do mês
mes_traduzido = {
    "JANUARY": "JANEIRO", "FEBRUARY": "FEVEREIRO", "MARCH": "MARÇO",
    "APRIL": "ABRIL", "MAY": "MAIO", "JUNE": "JUNHO",
    "JULY": "JULHO", "AUGUST": "AGOSTO", "SEPTEMBER": "SETEMBRO",
    "OCTOBER": "OUTUBRO", "NOVEMBER": "NOVEMBRO", "DECEMBER": "DEZEMBRO"
}.get(mes_nome_en, "MÊS_DESCONHECIDO")

# Caminhos dinâmicos
ARQUIVO_ORIGINAL = fr"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - {ano_atual}\{mes_num} - Fora de Rota automatico - {mes_traduzido}.xlsm"

PASTA_BACKUP = fr"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - {ano_atual}\BACKUPS\BACKUPS {mes_traduzido}"

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

    # Nome do backup
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