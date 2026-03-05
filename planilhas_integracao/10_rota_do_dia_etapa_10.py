# -*- coding: utf-8 -*-
"""
Etapa 09 - Rota do Dia
Copia os códigos da coluna A (linha 2+) do arquivo BASES_FILTRADA.csv
e cola como números na aba COORDENADAS, coluna B (a partir da linha 8)
do arquivo 10 - Fora de Rota automatico - OUTUBRO.xlsm
"""

import pandas as pd
import pythoncom
from win32com.client import DispatchEx
from pathlib import Path

# Caminhos fixos
XLSM_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
CSV_PATH = Path(r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\BASES_FILTRADA.csv")

# Aba e posição alvo
SHEET_NAME = "COORDENADAS"
COL_DEST = 2  # Coluna B
ROW_START = 8

def main():
    print("=== ETAPA 09 - ROTA DO DIA ===")

    # Verifica se CSV existe
    if not CSV_PATH.exists():
        print(f"❌ Arquivo CSV não encontrado:\n{CSV_PATH}")
        return

    # Lê CSV
    df = pd.read_csv(CSV_PATH, sep=";", encoding="utf-8-sig")
    if df.empty:
        print("⚠ O arquivo BASES_FILTRADA.csv está vazio.")
        return

    # Coluna A = Sold (remove valores nulos e garante int)
    codigos = (
        df.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.replace(r"\.0$", "", regex=True)
        .astype(int)
        .tolist()
    )
    print(f"📋 {len(codigos)} códigos prontos para colagem.")

    # Inicializa Excel com COM
    pythoncom.CoInitialize()
    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(XLSM_PATH)
        ws = wb.Worksheets(SHEET_NAME)

        # Limpa área destino (opcional)
        last_row = ROW_START + len(codigos) + 5
        ws.Range(ws.Cells(ROW_START, COL_DEST), ws.Cells(last_row, COL_DEST)).ClearContents()

        # Cola códigos como número
        for i, codigo in enumerate(codigos, start=ROW_START):
            ws.Cells(i, COL_DEST).Value = codigo

        print(f"✅ {len(codigos)} códigos colados com sucesso em {SHEET_NAME}!B{ROW_START}:")

        wb.Save()
        wb.Close(SaveChanges=True)
        print("💾 Alterações salvas com sucesso!")

    except Exception as e:
        print("❌ Erro durante execução:", e)

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
