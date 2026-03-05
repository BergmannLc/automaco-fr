# -*- coding: utf-8 -*-
"""
Etapa 11 - Delete Retidos (Padronizado)
Abre o arquivo 02 - Retidos.xlsx e exclui todas as linhas de 2 até 20000.
Nada além é alterado.
"""

import pythoncom
from win32com.client import DispatchEx
from pathlib import Path

# Caminho fixo da planilha
XLSX_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\02 - Retidos.xlsx"

def main():
    print("=== ETAPA 11 - DELETE RETIDOS (PADRÃO 2→20000) ===")

    # Verifica se o arquivo existe
    path = Path(XLSX_PATH)
    if not path.exists():
        print(f"❌ Arquivo não encontrado:\n{path}")
        return

    # Inicializa COM
    pythoncom.CoInitialize()
    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Abre a planilha
        wb = excel.Workbooks.Open(path.as_posix())
        ws = wb.Worksheets(1)  # Primeira aba (ajuste se precisar)

        # Exclui as linhas 2 a 20000
        print("🧹 Excluindo linhas de 2 até 20000...")
        ws.Range("2:20000").Delete()
        print("✅ Linhas 2 até 20000 excluídas com sucesso!")

        # Salva e fecha
        wb.Save()
        wb.Close(SaveChanges=True)
        print("💾 Alterações salvas com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao processar a planilha: {e}")

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
