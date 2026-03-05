# -*- coding: utf-8 -*-
"""
Etapa 16 - Colagem Resultado Final
Copia os valores da coluna N (Resultado_Final) do arquivo resultado_autorizacao_final.xlsx
e cola como valores na aba do dia vigente (coluna L, a partir da linha 3)
na planilha 10 - Fora de Rota automatico - OUTUBRO.xlsm.
"""

import pythoncom
from win32com.client import DispatchEx
import pandas as pd
from datetime import datetime
from pathlib import Path

# Caminhos fixos
SRC_RESULTADO = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\resultado_autorizacao_final.xlsx"
DST_PLANILHA = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"

def main():
    print("=== ETAPA 16 - COLAGEM RESULTADO FINAL ===")

    # 🗓 Identifica o dia vigente (aba a ser atualizada)
    dia_atual = datetime.now().day
    aba_dia = str(dia_atual)
    print(f"📅 Aba do dia vigente identificada: {aba_dia}")

    # 🔹 Lê os dados da planilha de resultado
    print("📂 Lendo resultado_autorizacao_final.xlsx...")
    try:
        df = pd.read_excel(SRC_RESULTADO, engine="openpyxl")
    except Exception as e:
        print(f"❌ Erro ao abrir {SRC_RESULTADO}: {e}")
        return

    # Verifica se a coluna N existe
    if "Resultado_Final" not in df.columns:
        print("❌ Coluna 'Resultado_Final' não encontrada no arquivo.")
        return

    # Pega apenas os valores da coluna N (ignorando o cabeçalho)
    valores = df["Resultado_Final"].dropna().tolist()
    total_valores = len(valores)
    print(f"📊 {total_valores} valores encontrados para colagem.")

    if total_valores == 0:
        print("⚠ Nenhum valor encontrado para colar. Abortando.")
        return

    # 🔹 Inicializa Excel (COM)
    pythoncom.CoInitialize()
    excel = DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        print("📂 Abrindo planilha principal...")
        wb = excel.Workbooks.Open(DST_PLANILHA)

        try:
            ws = wb.Worksheets(aba_dia)
        except Exception:
            print(f"❌ Aba '{aba_dia}' não encontrada.")
            wb.Close(SaveChanges=False)
            return

        # Define faixa de destino (coluna L, a partir da linha 3)
        linha_inicial = 3
        linha_final = linha_inicial + total_valores - 1
        print(f"📋 Colando de L{linha_inicial} até L{linha_final}...")

        ws.Range(f"L{linha_inicial}:L{linha_final}").Value = [[v] for v in valores]

        # Salva e fecha
        wb.Save()
        wb.Close(SaveChanges=True)
        print("✅ Colagem concluída com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao processar colagem: {e}")

    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
