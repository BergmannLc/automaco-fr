import subprocess
import os
import sys

# Lista organizada com o caminho correto de cada arquivo
fluxo_de_trabalho = [
    {"nome": "01_backup_etapa_1.py", "pasta": "planilhas_integracao"},
    {"nome": "02_atualizando_etapa_2.py", "pasta": "planilhas_integracao"},
    {"nome": "03_macro_etapa_3.py", "pasta": "planilhas_integracao"},
    {"nome": "04_copia_etapa_4.py", "pasta": "planilhas_integracao"},
    {"nome": "05_colagem_etapa_5.py", "pasta": "planilhas_integracao"},
    {"nome": "06_limpeza_etapa_6.py", "pasta": "planilhas_integracao"},
    {"nome": "07_colagem_coordenadas_etapa_7.py", "pasta": "planilhas_integracao"},
    {"nome": "09_filtragem_bases_etapa_9.py", "pasta": "planilhas_integracao"}, # Este pedirá sua interação
    {"nome": "10_rota_do_dia_etapa_10.py", "pasta": "planilhas_integracao"},
    {"nome": "11_delete_varejo_11.py", "pasta": "planilhas_integracao"},
    {"nome": "12_delete_retidos_12.py", "pasta": "planilhas_integracao"},
    {"nome": "13_delete_fora_de_rota_13.py", "pasta": "planilhas_integracao"},
    {"nome": "14_colagem_rota_14.py", "pasta": "planilhas_integracao"},
    {"nome": "16_colagem_fora_de_rota_16.py", "pasta": "planilhas_integracao"},
]

def executar_automacao():
    diretorio_raiz = os.getcwd()
    
    print("🚀 Iniciando Automação SAR...")
    print("-" * 30)

    for item in fluxo_de_trabalho:
        script = item["nome"]
        pasta_destino = os.path.join(diretorio_raiz, item["pasta"])
        caminho_completo = os.path.join(pasta_destino, script)

        if os.path.exists(caminho_completo):
            print(f"⏳ Executando: {script}...")
            
            # O segredo aqui é o 'cwd', que faz o Python rodar como se estivesse dentro da pasta do arquivo
            resultado = subprocess.run([sys.executable, caminho_completo], cwd=pasta_destino)

            if resultado.returncode == 0:
                print(f"✅ {script} concluído.\n")
            else:
                print(f"❌ Erro no arquivo {script}. Parando tudo para segurança.")
                return
        else:
            print(f"⚠️ Atenção: O arquivo {script} não foi encontrado em {pasta_destino}")
            return

    print("-" * 30)
    print("🏆 Todas as etapas foram concluídas com sucesso!")

if __name__ == "__main__":
    executar_automacao()