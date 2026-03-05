# -*- coding: utf-8 -*-
"""
Etapa 13 - Colagem Rota
Abre o arquivo 10 - Fora de Rota automatico - OUTUBRO.xlsm, lê a aba 'COORDENADAS'
e copia os valores das colunas B até I (a partir da linha 8 até a última linha com dados).
Depois, cola como número no arquivo 01 - Varejo.xlsx, na aba ativa, começando da linha 2, coluna A.
Nada além deve ser alterado.
"""

import pythoncom
from win32com.client import DispatchEx
from pathlib import Path

# Caminhos fixos das planilhas
SRC_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
DST_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\01 - Varejo.xlsx"

def main():
    print("=== ETAPA 13 - COLAGEM ROTA ===")

    # Inicializa COM
    pythoncom.CoInitialize()
    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Abre as duas planilhas
        print("📂 Abrindo arquivos...")
        wb_src = excel.Workbooks.Open(SRC_PATH)
        ws_src = wb_src.Worksheets("COORDENADAS")

        wb_dst = excel.Workbooks.Open(DST_PATH)
        ws_dst = wb_dst.Worksheets(1)  # Primeira aba do destino

        # Determina última linha da planilha origem (coluna B)
        last_row = ws_src.Cells(ws_src.Rows.Count, "B").End(-4162).Row  # xlUp = -4162
        if last_row < 8:
            print("⚠ Nenhum dado encontrado a partir da linha 8.")
            wb_src.Close(SaveChanges=False)
            wb_dst.Close(SaveChanges=False)
            return

        print(f"📊 Linhas encontradas: {last_row - 7}")

        # Copia da origem (colunas B até I)
        src_range = ws_src.Range(f"B8:I{last_row}")
        src_values = src_range.Value

        # Cola como número no destino (a partir da linha 2, coluna A)
        print("📋 Colando valores no destino...")
        dst_start = ws_dst.Cells(2, 1)
        dst_end = ws_dst.Cells(1 + len(src_values), 8)
        ws_dst.Range(dst_start, dst_end).Value = src_values

        # Salva e fecha
        wb_dst.Save()
        wb_dst.Close(SaveChanges=True)
        wb_src.Close(SaveChanges=False)

        print("✅ Colagem concluída com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao processar: {e}")

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
