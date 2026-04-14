import streamlit as st
import os
import shutil
from processador import rodar_analise

# -----------------------------
# PASTAS
# -----------------------------
UPLOAD_FOLDER = "uploads"
PROCESSADAS_FOLDER = "processadas"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSADAS_FOLDER, exist_ok=True)

st.set_page_config(page_title="IA Fotos", layout="centered")

st.title("📸 IA de Análise de Fotos")

# -----------------------------
# SESSION STATE (CHAVE DO FIX)
# -----------------------------
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# -----------------------------
# UPLOAD
# -----------------------------
uploaded_files = st.file_uploader(
    "Arraste ou selecione as imagens",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key=st.session_state.uploader_key
)

# -----------------------------
# BOTÕES
# -----------------------------
col1, col2 = st.columns(2)

# -----------------------------
# ANALISAR
# -----------------------------
if col1.button("🚀 Analisar"):

    if not uploaded_files:
        st.warning("⚠️ Envie pelo menos uma imagem")
    else:
        st.info("📥 Salvando imagens...")

        # limpa uploads
        for f in os.listdir(UPLOAD_FOLDER):
            os.remove(os.path.join(UPLOAD_FOLDER, f))

        # salva imagens
        for file in uploaded_files:
            with open(os.path.join(UPLOAD_FOLDER, file.name), "wb") as f:
                f.write(file.getbuffer())

        st.success("Imagens salvas!")

        with st.spinner("🧠 IA analisando..."):
            caminho_excel = rodar_analise(UPLOAD_FOLDER)

        st.success("✅ Análise concluída!")

        with open(caminho_excel, "rb") as f:
            st.download_button(
                "📥 Baixar Relatório",
                f,
                file_name="relatorio.xlsx"
            )

# -----------------------------
# RESET COMPLETO (AGORA FUNCIONA)
# -----------------------------
if col2.button("🗑️ Resetar"):

    # apagar pastas
    for pasta in [UPLOAD_FOLDER, PROCESSADAS_FOLDER]:
        if os.path.exists(pasta):
            shutil.rmtree(pasta)

    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(PROCESSADAS_FOLDER, exist_ok=True)

    # apagar excels
    for f in os.listdir():
        if f.endswith(".xlsx"):
            try:
                os.remove(f)
            except:
                pass

    # 🔥 RESET DO UPLOADER (ESSA É A CHAVE)
    st.session_state.uploader_key += 1

    st.success("🗑️ Sistema totalmente resetado!")

    st.rerun()