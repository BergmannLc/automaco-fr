# -*- coding: utf-8 -*-
"""
Filtragem BASES -> BASES_FILTRADA.csv
- Lê a planilha .xlsm original (sem corromper)
- Copia BASES!P:U
- Aplica filtros: Setor (whitelist), Dia, Período (sempre inclui '1234')
- Gera CSV com cabeçalho padronizado
- Corrigido: "Sold" sem .0 (força inteiro/texto)
"""

import re
import sys
import csv
import time
from pathlib import Path
import pythoncom
from win32com.client import DispatchEx

# ===== Caminho fixo do arquivo =====
XLSM_PATH = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Fora de rota - 2026\03 - Fora de Rota automatico - MARÇO.xlsm"

# ===== Setores permitidos =====
SETORES_OK = {
    118,119,121,150,151,152,153,154,201,114,123,133,134,140,
    301,307,308,401,402,403,404,405,406,407,501,502,503,504,
    505,506,602,603,604,605,606,801,802,803,804,805,806,807
}

# ===== Cabeçalho fixo =====
HEADER = ["Sold", "Razão", "Categoria", "Setor", "Dia", "Período"]

_re_int = re.compile(r"\d+")

def norm_int(v):
    """Extrai dígitos e retorna int, se houver."""
    if v is None:
        return None
    s = str(v).replace("\u00A0", " ").strip()
    m = _re_int.search(s)
    return int(m.group()) if m else None

def norm_str(v):
    return "" if v is None else str(v).replace("\u00A0", " ").strip()

def get_params():
    dia = input("Digite o dia da semana (1, 2, 3, 4, 5): ").strip()
    if dia not in {"1","2","3","4","5"}:
        print("Dia inválido."); sys.exit(1)
    periodo = input("Digite o Período, semanal sempre marcado (1 3 ou 2 4): ").strip()
    if periodo not in {"1 3","2 4"}:
        print("Período inválido."); sys.exit(1)
    return int(dia), periodo

def abrir_excel_com_retry(tentativas=8, espera_seg=1.0):
    """
    Abre uma instância nova do Excel via COM com retries.
    Evita:
      - erro de makepy do EnsureDispatch
      - 'A chamada foi rejeitada pelo chamado'
    """
    pythoncom.CoInitialize()
    last_err = None

    for _ in range(tentativas):
        try:
            xl = DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            return xl
        except Exception as e:
            last_err = e
            time.sleep(espera_seg)

    raise last_err

def main():
    print("=== FILTRO DE BASES (CSV com cabeçalho padronizado) ===")
    dia_alvo, periodo_alvo = get_params()
    print(f"\n🧮 Filtro -> Dia: {dia_alvo}, Período: {periodo_alvo} (+ '1234')\n")

    path = Path(XLSM_PATH)
    if not path.exists():
        print(f"❌ Arquivo não encontrado:\n{path}")
        sys.exit(1)

    xl = abrir_excel_com_retry()

    try:
        wb = xl.Workbooks.Open(path.as_posix())
        try:
            ws_src = wb.Worksheets("BASES")
        except:
            raise RuntimeError("Aba 'BASES' não encontrada.")

        # Força materialização do UsedRange
        ws_src.UsedRange
        last_row = ws_src.UsedRange.Rows.Count
        if last_row < 2:
            print("⚠ Nada a processar.")
            wb.Close(SaveChanges=False)
            return

        # Lê colunas P:U (16..21), incluindo cabeçalho na 1ª linha
        data = ws_src.Range(ws_src.Cells(1,16), ws_src.Cells(last_row,21)).Value
        rows = list(data[1:])  # ignora cabeçalho do Excel
        print(f"📋 Linhas lidas: {len(rows)}")

        # Índices dentro de P:U (0-based)
        IDX_SOLD, IDX_RAZ, IDX_CAT, IDX_VEND, IDX_DIA, IDX_PERI = 0, 1, 2, 3, 4, 5

        PERI_OK = {periodo_alvo, "1234"}

        kept = [HEADER]
        for r in rows:
            # Normalizações
            sold_int = norm_int(r[IDX_SOLD])      # força inteiro (sem .0)
            sold = "" if sold_int is None else str(sold_int)

            razao = norm_str(r[IDX_RAZ])
            categ = norm_str(r[IDX_CAT])
            vend = norm_int(r[IDX_VEND])
            dia = norm_int(r[IDX_DIA])
            peri = norm_str(r[IDX_PERI])

            # Filtros
            if (vend in SETORES_OK) and (dia == dia_alvo) and (peri in PERI_OK):
                kept.append([sold, razao, categ, vend, dia, peri])

        mantidas = len(kept) - 1
        removidas = len(rows) - mantidas
        print(f"✅ Linhas mantidas: {mantidas}")
        print(f"🗑️ Linhas removidas: {removidas}")

        # Escreve CSV com ; e BOM UTF-8
        csv_path = path.parent / "BASES_FILTRADA.csv"
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerows(kept)

        print(f"💾 CSV criado com sucesso: {csv_path}")

        wb.Close(SaveChanges=False)

    except Exception as e:
        print("❌ Erro:", e)
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
    finally:
        xl.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
