import os
import datetime
import pandas as pd
import numpy as np
from PIL import Image, ImageDraw
from ultralytics import YOLO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage

# -----------------------------
# CAMINHOS
# -----------------------------

PLANILHA = "planilha.xlsx"
MODELO = "best.pt"

DOWNLOADS = "."

agora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

RELATORIO = os.path.join(DOWNLOADS, f"relatorio_{agora}.xlsx")

PASTA_PROCESSADAS = "processadas"
os.makedirs(PASTA_PROCESSADAS, exist_ok=True)

# -----------------------------
# CARREGAR MODELO
# -----------------------------

model = YOLO(MODELO)

# -----------------------------
# PLANILHA
# -----------------------------

df = pd.read_excel(PLANILHA)

coluna_problema = "PROBLEMAS"
coluna_prioridade = "PRIORIDADE"

# -----------------------------
# FUNÇÕES
# -----------------------------

def ler_imagem(caminho):
    try:
        img = Image.open(caminho).convert("RGB")
        return img
    except:
        return None

def classificar_imagem(caminho_img):
    results = model(caminho_img)

    probs = results[0].probs

    if probs is None:
        return "Sem problemas", 0

    classe = results[0].names[probs.top1]
    confianca = float(probs.top1conf)

    if confianca < 0.30:
        return "Sem problemas", confianca

    return classe, confianca

# -----------------------------
# FUNÇÃO PRINCIPAL
# -----------------------------

def rodar_analise(PASTA_FOTOS):

    arquivos_imagem = []

    for raiz, _, arquivos in os.walk(PASTA_FOTOS):
        for arquivo in arquivos:
            if arquivo.lower().endswith((".jpg", ".jpeg", ".png")):
                arquivos_imagem.append(os.path.join(raiz, arquivo))

    print(f"Total de imagens: {len(arquivos_imagem)}")

    resultados = []

    for caminho in arquivos_imagem:

        nome = os.path.basename(caminho)

        imagem = ler_imagem(caminho)

        if imagem is None:
            continue

        categoria, conf = classificar_imagem(caminho)

        # desenhar texto na imagem (PIL)
        draw = ImageDraw.Draw(imagem)
        draw.text((10, 10), f"{categoria} ({conf:.2f})", fill=(255, 0, 0))

        caminho_saida = os.path.join(PASTA_PROCESSADAS, nome)
        imagem.save(caminho_saida)

        prioridade = ""

        if categoria != "Sem problemas":
            for _, row in df.iterrows():
                if categoria.lower() in str(row[coluna_problema]).lower():
                    prioridade = row[coluna_prioridade]
                    break

        resultados.append({
            "Foto": nome,
            "Categoria": categoria,
            "Prioridade": prioridade,
            "Confiança": round(conf, 3),
            "Imagem": caminho_saida
        })

    # -----------------------------
    # EXCEL
    # -----------------------------

    wb = Workbook()
    ws = wb.active
    ws.title = "Relatorio"

    ws.append(["Foto", "Categoria", "Prioridade", "Confiança", "Imagem"])

    linha = 2

    for r in resultados:

        ws.cell(linha, 1, r["Foto"])
        ws.cell(linha, 2, r["Categoria"])
        ws.cell(linha, 3, r["Prioridade"])
        ws.cell(linha, 4, r["Confiança"])

        if os.path.exists(r["Imagem"]):
            img = ExcelImage(r["Imagem"])
            img.width = 120
            img.height = 120
            ws.add_image(img, f"E{linha}")

        ws.row_dimensions[linha].height = 100

        linha += 1

    ws.column_dimensions["E"].width = 30

    # -----------------------------
    # RESUMO
    # -----------------------------

    ws_resumo = wb.create_sheet("Resumo")

    df_result = pd.DataFrame(resultados)

    # 🔥 IGNORA "Sem problemas"
    df_filtrado = df_result[df_result["Categoria"] != "Sem problemas"]

    media_prioridade = pd.to_numeric(df_filtrado["Prioridade"], errors='coerce').mean()

    emoji = "⚠️"

    if pd.notna(media_prioridade):
        if media_prioridade >= 4:
            emoji = "✅"
        elif media_prioridade <= 2:
            emoji = "🚨"

    ws_resumo.append(["Média Prioridade", f"{round(media_prioridade,2) if pd.notna(media_prioridade) else 'N/A'} {emoji}"])

    # contagem
    contagem = df_filtrado["Categoria"].value_counts()

    ws_resumo.append([])
    ws_resumo.append(["Problema", "Quantidade"])

    for problema, qtd in contagem.items():
        ws_resumo.append([problema, qtd])

    # -----------------------------
    wb.save(RELATORIO)

    return RELATORIO
