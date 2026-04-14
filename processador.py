import os
import datetime
import pandas as pd
import numpy as np
from PIL import Image
from transformers import CLIPProcessor, CLIPModel
from openpyxl import Workbook

# -----------------------------
# CONFIG
# -----------------------------

PLANILHA = "planilha.xlsx"

# -----------------------------
# CONHECIMENTO IA (OTIMIZADO)
# -----------------------------

conhecimento = {
"fiação exposta":"fios elétricos desencapados, perigosos e visíveis fora da parede",
"infiltrações":"parede com manchas de umidade, mofo, infiltração de água",
"rachaduras":"parede com trincas, fissuras ou rachaduras estruturais visíveis",
"pintura":"parede com tinta descascando, desgaste ou falhas na pintura",
"telhado":"telhas quebradas, deslocadas ou danificadas",
"calha entupida":"calha cheia de sujeira, folhas ou obstruída",
"revestimento":"revestimento quebrado ou soltando da parede",
"vidro":"vidro quebrado ou trincado",
"porta":"porta danificada, quebrada ou desalinhada",
"forro":"forro caindo, solto ou danificado no teto",
"torneira":"torneira vazando água",
"cano":"cano com vazamento ou dano",
"ralo":"ralo entupido ou sujo",
"cftv":"câmera de segurança instalada",
"porta de rolagem":"porta metálica enrolável",
"recarga extintor":"extintor de incêndio",
"sem_problema":"ambiente em bom estado, sem danos, sem problemas estruturais"
}

labels = list(conhecimento.keys())
textos = list(conhecimento.values())

superficies = {
"parede":"parede de alvenaria vertical",
"teto":"teto ou forro acima",
"chao":"piso ou chão"
}

labels_sup = list(superficies.keys())
textos_sup = list(superficies.values())

# -----------------------------
# MODELO CLIP
# -----------------------------

print("Carregando CLIP...")
clip_model = CLIPModel.from_pretrained("openai/clip-vit-base-patch32")
clip_processor = CLIPProcessor.from_pretrained("openai/clip-vit-base-patch32")
print("CLIP pronto!")

df = pd.read_excel(PLANILHA)

# -----------------------------
# FUNÇÕES IA
# -----------------------------

def analisar_clip(imagem):

    inputs = clip_processor(
        text=textos,
        images=imagem,
        return_tensors="pt",
        padding=True
    )

    outputs = clip_model(**inputs)
    probs = outputs.logits_per_image.softmax(dim=1)[0]

    top = probs.topk(3)

    resultados = []

    for i in range(3):
        idx = top.indices[i].item()
        conf = float(top.values[i])
        resultados.append((labels[idx], conf))

    return resultados


def detectar_superficie(imagem):

    inputs = clip_processor(
        text=textos_sup,
        images=imagem,
        return_tensors="pt",
        padding=True
    )

    outputs = clip_model(**inputs)
    probs = outputs.logits_per_image.softmax(dim=1)[0]

    idx = probs.argmax().item()

    return labels_sup[idx]

# -----------------------------
# MULTI-ANÁLISE (SEM CV2)
# -----------------------------

def gerar_cortes(img_np):

    h, w, _ = img_np.shape

    cortes = []

    # imagem inteira
    cortes.append(img_np)

    # centro
    cortes.append(img_np[h//4:3*h//4, w//4:3*w//4])

    # canto superior
    cortes.append(img_np[0:h//2, 0:w//2])

    # canto inferior
    cortes.append(img_np[h//2:h, w//2:w])

    return cortes

# -----------------------------
# FILTRO INTELIGENTE
# -----------------------------

def escolher(resultado, superficie):

    for categoria, conf in resultado:

        if conf < 0.35:
            continue

        if categoria == "sem_problema":
            return "Sem problemas", conf

        if categoria in ["rachaduras","infiltrações"] and superficie != "parede":
            continue

        if categoria in ["torneira","ralo","cano"] and superficie == "teto":
            continue

        return categoria, conf

    return "Sem problemas", 0

# -----------------------------
# MAIN
# -----------------------------

def rodar_analise(pasta):

    resultados = []

    for arquivo in os.listdir(pasta):

        if not arquivo.lower().endswith((".jpg",".png",".jpeg")):
            continue

        caminho = os.path.join(pasta, arquivo)

        print("Analisando:", arquivo)

        try:
            img_pil_full = Image.open(caminho).convert("RGB")
        except:
            continue

        img_np = np.array(img_pil_full)

        cortes = gerar_cortes(img_np)

        votos = []

        for corte in cortes:

            img_pil = Image.fromarray(corte)

            superficie = detectar_superficie(img_pil)
            res = analisar_clip(img_pil)

            cat, conf = escolher(res, superficie)

            votos.append((cat, conf))

        # votação final
        categorias = [v[0] for v in votos]
        categoria_final = max(set(categorias), key=categorias.count)

        conf_final = max([v[1] for v in votos if v[0] == categoria_final], default=0)

        prioridade = ""

        if categoria_final != "Sem problemas":
            for _, row in df.iterrows():
                if categoria_final.lower() in str(row["PROBLEMAS"]).lower():
                    prioridade = row["PRIORIDADE"]
                    break

        resultados.append({
            "Foto": arquivo,
            "Categoria": categoria_final,
            "Prioridade": prioridade,
            "Confiança": round(conf_final,3)
        })

    # -----------------------------
    # EXCEL
    # -----------------------------

    agora = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    nome = f"relatorio_{agora}.xlsx"

    wb = Workbook()
    ws = wb.active

    ws.append(["Foto","Categoria","Prioridade","Confiança"])

    for r in resultados:
        ws.append(list(r.values()))

    # resumo
    ws2 = wb.create_sheet("Resumo")

    df_result = pd.DataFrame(resultados)
    df_filtrado = df_result[df_result["Categoria"] != "Sem problemas"]

    media = pd.to_numeric(df_filtrado["Prioridade"], errors='coerce').mean()

    emoji = "⚠️"

    if pd.notna(media):
        if media <= 2:
            emoji = "🚨"
        elif media >= 4:
            emoji = "✅"

    ws2.append(["Média Prioridade", f"{round(media,2) if pd.notna(media) else 'N/A'} {emoji}"])

    contagem = df_filtrado["Categoria"].value_counts()

    ws2.append([])
    ws2.append(["Problema","Qtd"])

    for p, q in contagem.items():
        ws2.append([p,q])

    wb.save(nome)

    return nome
