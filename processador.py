import os
import datetime
import pandas as pd
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

# -----------------------------
# CONFIG
# -----------------------------

PLANILHA = "planilha.xlsx"
PASTA_PROCESSADAS = "processadas"

os.makedirs(PASTA_PROCESSADAS, exist_ok=True)

agora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RELATORIO = f"relatorio_{agora}.xlsx"

df = pd.read_excel(PLANILHA)

coluna_problema = "PROBLEMAS"
coluna_prioridade = "PRIORIDADE"

# -----------------------------
# CLASSIFICAÇÃO SIMPLES (FAKE IA)
# -----------------------------

def classificar_imagem(nome_arquivo):
    
    nome = nome_arquivo.lower()

    if "rachadura" in nome:
        return "rachaduras", 0.9
    elif "infiltracao" in nome:
        return "infiltrações", 0.85
    elif "fio" in nome:
        return "fiação exposta", 0.95

    return "Sem problemas", 0.5

# -----------------------------
# MAIN
# -----------------------------

def rodar_analise(PASTA_FOTOS):

    resultados = []

    for arquivo in os.listdir(PASTA_FOTOS):

        if not arquivo.lower().endswith((".jpg",".png",".jpeg")):
            continue

        caminho = os.path.join(PASTA_FOTOS, arquivo)

        categoria, conf = classificar_imagem(arquivo)

        prioridade = ""

        if categoria != "Sem problemas":
            for _, row in df.iterrows():
                if categoria.lower() in str(row[coluna_problema]).lower():
                    prioridade = row[coluna_prioridade]
                    break

        resultados.append({
            "Foto": arquivo,
            "Categoria": categoria,
            "Prioridade": prioridade,
            "Confiança": conf,
            "Imagem": caminho
        })

    # -----------------------------
    # EXCEL
    # -----------------------------

    wb = Workbook()
    ws = wb.active

    ws.append(["Foto","Categoria","Prioridade","Confiança"])

    for r in resultados:
        ws.append([r["Foto"], r["Categoria"], r["Prioridade"], r["Confiança"]])

    # -----------------------------
    # RESUMO
    # -----------------------------

    ws2 = wb.create_sheet("Resumo")

    df_result = pd.DataFrame(resultados)

    df_filtrado = df_result[df_result["Categoria"] != "Sem problemas"]

    media = pd.to_numeric(df_filtrado["Prioridade"], errors='coerce').mean()

    emoji = "⚠️"
    if pd.notna(media):
        if media >= 4:
            emoji = "✅"
        elif media <= 2:
            emoji = "🚨"

    ws2.append(["Média Prioridade", f"{media} {emoji}"])

    wb.save(RELATORIO)

    return RELATORIO
