# -*- coding: utf-8 -*-
"""
Etapa 14 - Colagem Retidos
Abre o arquivo 10 - Fora de Rota automatico - OUTUBRO.xlsm (aba COORDENADAS)
e copia os valores das colunas K até Q (linhas 8 até o final).
Depois, cola como número na planilha 02 - Retidos.xlsx a partir da linha 2 da coluna A.
Nada além deve ser alterado.
"""

import pythoncom
from win32com.client import DispatchEx
from pathlib import Path

# Caminhos fixos das planilhas
SRC_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"
DST_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\02 - Retidos.xlsx"

def main():
    print("=== ETAPA 14 - COLAGEM RETIDOS ===")

    pythoncom.CoInitialize()
    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Abre origem e destino
        print("📂 Abrindo arquivos...")
        wb_src = excel.Workbooks.Open(SRC_PATH)
        ws_src = wb_src.Worksheets("COORDENADAS")

        wb_dst = excel.Workbooks.Open(DST_PATH)
        ws_dst = wb_dst.Worksheets(1)  # Primeira aba

        # Identifica última linha da origem (base na coluna K)
        last_row = ws_src.Cells(ws_src.Rows.Count, "K").End(-4162).Row  # xlUp
        if last_row < 8:
            print("⚠ Nenhum dado encontrado a partir da linha 8.")
            wb_src.Close(SaveChanges=False)
            wb_dst.Close(SaveChanges=False)
            return

        total_rows = last_row - 7
        print(f"📊 Linhas encontradas: {total_rows}")

        # Copia colunas K:Q da origem
        src_range = ws_src.Range(f"K8:Q{last_row}")
        src_values = src_range.Value

        # Cola os valores no destino (A2)
        print("📋 Colando valores na planilha de destino...")
        dst_start = ws_dst.Cells(2, 1)
        dst_end = ws_dst.Cells(1 + len(src_values), 7)
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
