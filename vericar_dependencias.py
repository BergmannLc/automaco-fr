import os
import sys
import importlib
import subprocess


if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


DEPENDENCIAS_PIP = [
    "pywin32",
    "openpyxl",
    "pandas",
    "pyperclip",
    "selenium",
]


MODULOS_TESTE = [
    "win32com.client",
    "pythoncom",
    "openpyxl",
    "pandas",
    "pyperclip",
    "selenium",
]


def instalar_pacote(pacote):
    print(f"\nInstalando pacote: {pacote}...")

    resultado = subprocess.run(
        [sys.executable, "-m", "pip", "install", pacote],
        capture_output=True,
        text=True
    )

    if resultado.returncode != 0:
        print(f"\nErro ao instalar {pacote}")
        print(resultado.stderr)
        return False

    print(f"Pacote {pacote} instalado com sucesso.")
    return True


def verificar_e_instalar_dependencias():
    print("\nVerificando dependências...")

    for pacote, modulo in zip(DEPENDENCIAS_PIP, MODULOS_TESTE):
        try:
            importlib.import_module(modulo)
            print(f"OK: {modulo}")
        except ImportError:
            print(f"FALTANDO: {modulo}")
            sucesso = instalar_pacote(pacote)

            if not sucesso:
                print("\nFalha crítica na instalação. Encerrando.")
                input("Pressione ENTER para sair...")
                sys.exit(1)

            # tenta importar novamente após instalar
            try:
                importlib.import_module(modulo)
                print(f"OK após instalação: {modulo}")
            except ImportError:
                print(f"\nErro ao importar {modulo} mesmo após instalação.")
                input("Pressione ENTER para sair...")
                sys.exit(1)


def verificar_arquivos():
    arquivos = [
        "chromedriver.exe",
        "Vendas_Do_Dia.xlsx",
        "debug_vendas_fora_rota.txt",
    ]

    for arquivo in arquivos:
        caminho = os.path.join(BASE_DIR, arquivo)
        if not os.path.exists(caminho):
            print(f"\nArquivo obrigatório não encontrado: {arquivo}")
            input("Pressione ENTER para sair...")
            sys.exit(1)


def verificar_pastas():
    pastas = [
        "planilhas_integracao",
        "ChromeProfile",
        "ChromeTemp",
    ]

    for pasta in pastas:
        caminho = os.path.join(BASE_DIR, pasta)
        if not os.path.exists(caminho):
            print(f"\nPasta obrigatória não encontrada: {pasta}")
            input("Pressione ENTER para sair...")
            sys.exit(1)


def main():
    print("\n" + "=" * 80)
    print("VERIFICAÇÃO AUTOMÁTICA DE DEPENDÊNCIAS")
    print("=" * 80)

    verificar_e_instalar_dependencias()
    verificar_arquivos()
    verificar_pastas()

    print("\nTudo pronto para execução.")
    print("=" * 80)


if __name__ == "__main__":
    main()