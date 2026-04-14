import os
import sys
import shutil
import subprocess


if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PASTA_PLANILHAS = os.path.join(BASE_DIR, "planilhas_integracao")


def limpar_tela():
    os.system("cls" if os.name == "nt" else "clear")


def pausar():
    input("\nPressione ENTER para continuar...")


def obter_comando_python():
    if not getattr(sys, "frozen", False):
        return [sys.executable]

    python_cmd = shutil.which("python")
    if python_cmd:
        return [python_cmd]

    py_launcher = shutil.which("py")
    if py_launcher:
        return [py_launcher, "-3"]

    raise RuntimeError(
        "Não foi encontrado um interpretador Python instalado neste computador.\n"
        "Instale o Python e tente novamente."
    )


def rodar_script(caminho_script, nome_exibicao=None):
    if nome_exibicao is None:
        nome_exibicao = os.path.basename(caminho_script)

    print("\n" + "=" * 80)
    print(f"INICIANDO: {nome_exibicao}")
    print("=" * 80)

    if not os.path.exists(caminho_script):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_script}")

    comando = obter_comando_python() + [caminho_script]

    resultado = subprocess.run(
        comando,
        cwd=BASE_DIR
    )

    if resultado.returncode != 0:
        raise RuntimeError(
            f"O script '{nome_exibicao}' terminou com erro (código {resultado.returncode})."
        )

    print(f"\nFINALIZADO COM SUCESSO: {nome_exibicao}")


def verificar_dependencias_iniciais():
    rodar_script(
        os.path.join(BASE_DIR, "verificar_dependencias.py"),
        "verificar_dependencias.py"
    )


def perguntar_retidos():
    while True:
        resposta = input(
            "\nO txt foi atualizado com os retidos de hoje? (S/N): "
        ).strip().upper()

        if resposta == "S":
            return "executar"

        elif resposta == "N":
            while True:
                resposta2 = input(
                    "\nDigite 'Aguardar' para aguardar o txt ser preenchido "
                    "ou 'Continuar' para continuar sem retidos: "
                ).strip().lower()

                if resposta2 == "aguardar":
                    input("\nQuando preencher o txt, aperte ENTER.")
                    return "executar"

                elif resposta2 == "continuar":
                    return "pular"

                else:
                    print("\nResposta inválida. Digite apenas 'Aguardar' ou 'Continuar'.")

        else:
            print("\nResposta inválida. Digite apenas S ou N.")


def pausa_conferencia_final():
    print("\n" + "=" * 80)
    print("PAUSA OBRIGATÓRIA PARA CONFERÊNCIA FINAL")
    print("=" * 80)
    print("1. Verifique as devolutivas geradas pelo código.")
    print("2. Confira a planilha.")
    print("3. Confira o sistema / mapa.")
    print("4. Só continue quando tiver certeza de que está tudo correto.")

    while True:
        resposta = input("\nDigite 'CONTINUAR' para seguir com o envio: ").strip().upper()

        if resposta == "CONTINUAR":
            return

        print("\nResposta inválida. Digite exatamente CONTINUAR.")


def main():
    try:
        limpar_tela()
        print("=" * 80)
        print("AUTOMAÇÃO SAR - EXECUÇÃO COMPLETA")
        print("=" * 80)
        print(f"\nPasta base identificada:\n{BASE_DIR}")

        verificar_dependencias_iniciais()

        scripts_iniciais = [
            os.path.join(BASE_DIR, "Novo automatizar_vendas.py"),
            os.path.join(PASTA_PLANILHAS, "01_backup_etapa_1.py"),
            os.path.join(PASTA_PLANILHAS, "02_atualizando_etapa_2.py"),
            os.path.join(PASTA_PLANILHAS, "03_macro_etapa_3.py"),
            os.path.join(PASTA_PLANILHAS, "04_copia_etapa_4.py"),
            os.path.join(PASTA_PLANILHAS, "05_colagem_etapa_5.py"),
            os.path.join(PASTA_PLANILHAS, "06_limpeza_etapa_6.py"),
            os.path.join(PASTA_PLANILHAS, "07_colagem_coordenadas_etapa_7.py"),
        ]

        for script in scripts_iniciais:
            rodar_script(script)

        acao_retidos = perguntar_retidos()

        if acao_retidos == "executar":
            rodar_script(os.path.join(PASTA_PLANILHAS, "08_retidos_etapa_8.py"))
        else:
            print("\n08_retidos_etapa_8.py foi pulado por escolha do usuário.")

        scripts_finais = [
            os.path.join(PASTA_PLANILHAS, "09_filtragem_bases_etapa_9.py"),
            os.path.join(PASTA_PLANILHAS, "10_rota_do_dia_etapa_10.py"),
            os.path.join(PASTA_PLANILHAS, "11_delete_varejo_11.py"),
            os.path.join(PASTA_PLANILHAS, "12_delete_retidos_12.py"),
            os.path.join(PASTA_PLANILHAS, "13_delete_fora_de_rota_13.py"),
            os.path.join(PASTA_PLANILHAS, "14_colagem_rota_14.py"),
            os.path.join(PASTA_PLANILHAS, "15_colagem_retidos_15.py"),
            os.path.join(PASTA_PLANILHAS, "16_colagem_fora_de_rota_16.py"),
            os.path.join(BASE_DIR, "analise_fora_de_rota.py"),
            os.path.join(PASTA_PLANILHAS, "17_colagem_resultado_final_17.py"),
        ]

        for script in scripts_finais:
            rodar_script(script)

        pausa_conferencia_final()

        rodar_script(os.path.join(BASE_DIR, "envio_forarota.py"))

        print("\n" + "=" * 80)
        print("PROCESSO FINALIZADO COM SUCESSO")
        print("=" * 80)
        pausar()

    except Exception as e:
        print("\n" + "=" * 80)
        print("ERRO NA EXECUÇÃO")
        print("=" * 80)
        print(f"\nDetalhes: {e}")
        pausar()


if __name__ == "__main__":
    main()