#Análise de fora de rota

import pandas as pd
import os
import math

# 🔹 Função Haversine (distância em km)
def haversine(lat1, lon1, lat2, lon2):
    R = 6371  # Raio da Terra em km
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    c = 2 * math.asin(math.sqrt(a))
    return R * c

# 🔹 Função para separar latitude e longitude
def separar_lat_lon(df, coluna):
    if coluna in df.columns:
        # Verifica se o DF não está vazio antes de tentar splitar
        if not df.empty:
            df[["Latitude", "Longitude"]] = df[coluna].astype(str).str.split(",", expand=True)
            df["Latitude"] = df["Latitude"].str.strip().astype(float)
            df["Longitude"] = df["Longitude"].str.strip().astype(float)
        else:
            df["Latitude"] = pd.Series(dtype=float)
            df["Longitude"] = pd.Series(dtype=float)
    return df

# 🔹 Caminho da pasta
caminho = r"\\192.168.254.64\Grupo Fast\SAR\6. Fora de Rota\Google Maps"

# Arquivos
arquivo_rota = os.path.join(caminho, "01 - Varejo.xlsx")
arquivo_fora = os.path.join(caminho, "03 - Fora de Rota.xlsx")
arquivo_retidos = os.path.join(caminho, "02 - Retidos.xlsx")

# 1️⃣ Primeira verificação (fora de rota vs varejo)
rota = pd.read_excel(arquivo_rota, engine="openpyxl")
fora = pd.read_excel(arquivo_fora, engine="openpyxl")

rota = separar_lat_lon(rota, "Coordenadas")
fora = separar_lat_lon(fora, "Coordenadas")

status = []
distancias_min = []
for i, row in fora.iterrows():
    lat_f, lon_f = row["Latitude"], row["Longitude"]
    distancias = rota.apply(lambda r: haversine(lat_f, lon_f, r["Latitude"], r["Longitude"]), axis=1)
    menor_dist = distancias.min()
    distancias_min.append(round(menor_dist, 2))
    status.append("Autorizado" if menor_dist <= 6 else "Não Autorizado")

fora["Status_Varejo"] = status
fora["Distancia_Varejo_km"] = distancias_min

# 2️⃣ Segunda verificação (apenas os não autorizados vs retidos)
retidos = pd.read_excel(arquivo_retidos, engine="openpyxl")

# --- VERIFICAÇÃO SE RETIDOS ESTÁ VAZIO ---
if retidos.empty:
    print("\n⚠️  AVISO: Planilha de Retidos está vazia. A análise foi feita apenas com base nos clientes de Varejo.")
    # Se estiver vazio, criamos as colunas vazias para não quebrar o merge
    fora["Status_Retidos"] = "Não Aplicável (Retidos Vazio)"
    fora["Distancia_Retidos_km"] = 0.0
else:
    retidos = separar_lat_lon(retidos, "Coordenadas")
    nao_autorizados = fora[fora["Status_Varejo"] == "Não Autorizado"].copy()

    status_retidos = []
    distancias_retidos = []
    
    for i, row in nao_autorizados.iterrows():
        lat_f, lon_f = row["Latitude"], row["Longitude"]
        distancias = retidos.apply(lambda r: haversine(lat_f, lon_f, r["Latitude"], r["Longitude"]), axis=1)
        menor_dist = distancias.min()
        distancias_retidos.append(round(menor_dist, 2))
        status_retidos.append("Autorizado_Retidos" if menor_dist <= 6 else "Nao_Autorizado_Final")

    nao_autorizados["Status_Retidos"] = status_retidos
    nao_autorizados["Distancia_Retidos_km"] = distancias_retidos

    # 3️⃣ Juntar apenas se houve análise de retidos
    fora = fora.merge(
        nao_autorizados[["Latitude", "Longitude", "Status_Retidos", "Distancia_Retidos_km"]],
        on=["Latitude", "Longitude"],
        how="left"
    )

# 🔹 Criar coluna final com o resultado definitivo
def definir_resultado(row):
    if row["Status_Varejo"] == "Autorizado":
        return "Autorizado"
    elif row["Status_Varejo"] == "Não Autorizado":
        # Se veio de retidos e deu autorizado
        if row.get("Status_Retidos") == "Autorizado_Retidos":
            return "Autorizado"
        else:
            return "Não Autorizado"
    else:
        return "Indefinido"

fora["Resultado_Final"] = fora.apply(definir_resultado, axis=1)

# 🔹 Salvar arquivo final
resultado_final = os.path.join(caminho, "resultado_autorizacao_final.xlsx")
fora.to_excel(resultado_final, index=False)

print(f"✅ Análise concluída! Resultado salvo em: {resultado_final}")
print("📊 Coluna 'Resultado_Final' adicionada com sucesso!")

# 🔹 Resumo rápido no terminal
resumo = fora["Resultado_Final"].value_counts()
print("\n📈 RESUMO FINAL:")
print(resumo)