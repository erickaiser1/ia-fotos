import os
import datetime
import pandas as pd
import cv2
from ultralytics import YOLO
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.chart import BarChart, Reference

# -----------------------------
# CONFIG
# -----------------------------

PLANILHA = "planilha.xlsx"
MODELO = "best.pt"  # modelo YOLO

DOWNLOADS = os.getcwd()

agora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
RELATORIO = os.path.join(DOWNLOADS, f"relatorio_{agora}.xlsx")

PASTA_PROCESSADAS = os.path.join(DOWNLOADS, "processadas")
os.makedirs(PASTA_PROCESSADAS, exist_ok=True)

# -----------------------------
# PLANILHA
# -----------------------------

df = pd.read_excel(PLANILHA)

coluna_problema = "PROBLEMAS"
coluna_prioridade = "PRIORIDADE"

# -----------------------------
# MODELO
# -----------------------------

print("Carregando modelo YOLO...")
model = YOLO(MODELO)
print("Modelo carregado!")

# -----------------------------
# CLASSIFICAÇÃO
# -----------------------------

def classificar_imagem(caminho_img):

    results = model(caminho_img)

    probs = results[0].probs

    if probs is None:
        return "Sem problemas", 0

    classe = results[0].names[probs.top1]
    conf = float(probs.top1conf)

    # regra mínima de confiança
    if conf < 0.30:
        return "Sem problemas", conf

    return classe, conf

# -----------------------------
# MAIN
# -----------------------------

def rodar_analise(PASTA_FOTOS):

    resultados = []

    for arquivo in os.listdir(PASTA_FOTOS):

        if not arquivo.lower().endswith((".jpg",".png",".jpeg")):
            continue

        caminho = os.path.join(PASTA_FOTOS, arquivo)

        print("Analisando:", arquivo)

        imagem = cv2.imread(caminho)

        if imagem is None:
            continue

        categoria, conf = classificar_imagem(caminho)

        # desenhar texto
        cv2.putText(imagem, f"{categoria} ({conf:.2f})",
                    (20,40),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    1,(0,0,255),2)

        caminho_saida = os.path.join(PASTA_PROCESSADAS, arquivo)
        cv2.imwrite(caminho_saida, imagem)

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
            "Confiança": round(conf,3),
            "Imagem": caminho_saida
        })

    # -----------------------------
    # EXCEL
    # -----------------------------

    wb = Workbook()
    ws = wb.active
    ws.title = "Relatorio"

    ws.append(["Foto","Categoria","Prioridade","Confiança","Imagem"])

    linha = 2

    for r in resultados:

        ws.cell(linha,1,r["Foto"])
        ws.cell(linha,2,r["Categoria"])
        ws.cell(linha,3,r["Prioridade"])
        ws.cell(linha,4,r["Confiança"])

        img = ExcelImage(r["Imagem"])
        img.width = 150
        img.height = 150

        ws.add_image(img,f"E{linha}")
        ws.row_dimensions[linha].height = 120

        linha += 1

    ws.column_dimensions["E"].width = 40

    # -----------------------------
    # RESUMO (SEM GRÁFICO)
    # -----------------------------

    df_result = pd.DataFrame(resultados)

    ws_resumo = wb.create_sheet("Resumo")

    df_filtrado = df_result[df_result["Categoria"] != "Sem problemas"]

    media = pd.to_numeric(df_filtrado["Prioridade"], errors='coerce').mean()

    emoji = "⚠️"

    if pd.notna(media):
        if media <= 2:
            emoji = "🚨"
        elif media >= 4:
            emoji = "✅"

    ws_resumo.append(["Média Prioridade", f"{round(media,2) if pd.notna(media) else 'N/A'} {emoji}"])

    # contagem (sem "Sem problemas")
    contagem = df_filtrado["Categoria"].value_counts()

    ws_resumo.append([])
    ws_resumo.append(["Problema","Quantidade"])

    for problema, qtd in contagem.items():
        ws_resumo.append([problema,qtd])

    # -----------------------------

    wb.save(RELATORIO)

    print("\n✅ Relatório criado:")
    print(RELATORIO)

    return RELATORIO
