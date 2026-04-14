# -*- coding: utf-8 -*-
"""
Etapa 16 - Colagem Resultado Final via PROCX

Nova lógica:
1. Abre a planilha resultado_autorizacao_final.xlsx
2. Abre a planilha mensal do Fora de Rota
3. Vai na aba do dia vigente
4. Localiza a coluna "Sold"
5. Insere fórmula PROCX na coluna L, a partir da linha 3
6. Preenche a fórmula até a última linha com dados
7. Converte a coluna L de fórmula para valores
"""

import os
import time
import pythoncom
from win32com.client import DispatchEx
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES DINÂMICAS
# ============================================================
agora = datetime.now()
ano_atual = agora.year
mes_num = agora.strftime("%m")
mes_nome_en = agora.strftime("%B").upper()

mes_traduzido = {
    "JANUARY": "JANEIRO",
    "FEBRUARY": "FEVEREIRO",
    "MARCH": "MARÇO",
    "APRIL": "ABRIL",
    "MAY": "MAIO",
    "JUNE": "JUNHO",
    "JULY": "JULHO",
    "AUGUST": "AGOSTO",
    "SEPTEMBER": "SETEMBRO",
    "OCTOBER": "OUTUBRO",
    "NOVEMBER": "NOVEMBRO",
    "DECEMBER": "DEZEMBRO"
}.get(mes_nome_en, "MÊS_DESCONHECIDO")

SRC_RESULTADO = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps\resultado_autorizacao_final.xlsx"
DST_PLANILHA = fr"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - {ano_atual}\{mes_num} - Fora de Rota automatico - {mes_traduzido}.xlsm"

# ============================================================
# AUXILIARES
# ============================================================
def abrir_excel_com_retry(tentativas=8, espera_seg=1.0):
    pythoncom.CoInitialize()
    last_err = None

    for _ in range(tentativas):
        try:
            excel = DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            return excel
        except Exception as e:
            last_err = e
            time.sleep(espera_seg)

    raise last_err


def numero_para_letra_coluna(col_num):
    resultado = ""
    while col_num > 0:
        col_num, resto = divmod(col_num - 1, 26)
        resultado = chr(65 + resto) + resultado
    return resultado


def localizar_coluna_por_cabecalho(ws, nome_cabecalho, linha_cabecalho=2, max_colunas=100):
    nome_procurado = str(nome_cabecalho).strip().lower()

    for col in range(1, max_colunas + 1):
        valor = ws.Cells(linha_cabecalho, col).Value
        if valor is None:
            continue

        if str(valor).strip().lower() == nome_procurado:
            return col

    return None


def ultima_linha_com_dados_na_coluna(ws, coluna, linha_minima=3):
    ultima = ws.Cells(ws.Rows.Count, coluna).End(-4162).Row  # xlUp = -4162
    if ultima < linha_minima:
        return linha_minima - 1
    return ultima

# ============================================================
# PRINCIPAL
# ============================================================
def main():
    print("=== ETAPA 16 - COLAGEM RESULTADO FINAL VIA PROCX ===")

    if not os.path.exists(SRC_RESULTADO):
        print(f"❌ Arquivo de origem não encontrado: {SRC_RESULTADO}")
        return

    if not os.path.exists(DST_PLANILHA):
        print(f"❌ Arquivo de destino não encontrado: {DST_PLANILHA}")
        return

    dia_atual = datetime.now().day
    aba_dia = str(dia_atual)
    print(f"📅 Aba do dia vigente identificada: {aba_dia}")

    excel = abrir_excel_com_retry()

    wb_resultado = None
    wb_destino = None

    try:
        print("📂 Abrindo arquivo de resultado...")
        wb_resultado = excel.Workbooks.Open(SRC_RESULTADO)

        print("📂 Abrindo planilha principal...")
        wb_destino = excel.Workbooks.Open(DST_PLANILHA)

        try:
            ws_dia = wb_destino.Worksheets(aba_dia)
        except Exception:
            print(f"❌ Aba '{aba_dia}' não encontrada na planilha destino.")
            return

        # Localiza a coluna "Sold" na linha 2
        col_sold = localizar_coluna_por_cabecalho(ws_dia, "Sold", linha_cabecalho=2, max_colunas=100)
        if not col_sold:
            print("❌ Cabeçalho 'Sold' não encontrado na linha 2 da aba do dia.")
            return

        letra_col_sold = numero_para_letra_coluna(col_sold)
        print(f"✅ Coluna 'Sold' localizada em: {letra_col_sold}")

        # Define intervalo com base na última linha preenchida da coluna Sold
        linha_inicial = 3
        linha_final = ultima_linha_com_dados_na_coluna(ws_dia, col_sold, linha_minima=linha_inicial)

        if linha_final < linha_inicial:
            print("⚠ Nenhum dado encontrado abaixo do cabeçalho da coluna 'Sold'.")
            return

        print(f"📋 Aplicando PROCX de L{linha_inicial} até L{linha_final}...")

        # Insere fórmula na primeira linha
        # Exemplo gerado:
        # =PROCX(C3;[resultado_autorizacao_final.xlsx]Sheet1!$A:$A;[resultado_autorizacao_final.xlsx]Sheet1!$N:$N;0;0)
        formula_linha_inicial = (
            f"=PROCX({letra_col_sold}{linha_inicial};"
            f"[resultado_autorizacao_final.xlsx]Sheet1!$A:$A;"
            f"[resultado_autorizacao_final.xlsx]Sheet1!$N:$N;"
            f"0;0)"
        )

        ws_dia.Range(f"L{linha_inicial}").FormulaLocal = formula_linha_inicial

        # Preenche até a última linha
        if linha_final > linha_inicial:
            ws_dia.Range(f"L{linha_inicial}").AutoFill(
                Destination=ws_dia.Range(f"L{linha_inicial}:L{linha_final}")
            )

        # Converte fórmulas em valores
        print("🔄 Convertendo fórmulas em valores...")
        rng = ws_dia.Range(f"L{linha_inicial}:L{linha_final}")
        rng.Value = rng.Value

        wb_destino.Save()
        print("✅ Processo concluído com sucesso!")

    except Exception as e:
        print(f"❌ Erro ao processar: {e}")

    finally:
        try:
            if wb_resultado is not None:
                wb_resultado.Close(SaveChanges=False)
        except Exception:
            pass

        try:
            if wb_destino is not None:
                wb_destino.Close(SaveChanges=True)
        except Exception:
            pass

        try:
            excel.Quit()
        except Exception:
            pass

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()