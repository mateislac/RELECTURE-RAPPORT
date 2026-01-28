import streamlit as st
from docx import Document
from openai import OpenAI
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import tempfile

# ======================
# MOT DE PASSE
# ======================
PASSWORD = "MALAC"

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("🔐 Accès sécurisé")

    pwd = st.text_input("Mot de passe", type="password")

    if pwd:
        if pwd == PASSWORD:
            st.session_state.auth = True
            st.success("Accès autorisé ✅")
            st.rerun()
        else:
            st.error("Mot de passe incorrect ❌")

    st.stop()

# ======================
# CONFIG OPENAI
# ======================
if not os.getenv("OPENAI_API_KEY"):
    st.error("❌ Clé API OpenAI manquante")
    st.stop()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

SYSTEM_PROMPT = """
Tu es un correcteur professionnel français pour rapports de bureau d’étude géotechnique.
Tu corriges l’orthographe, la grammaire et la formulation.
Tu reformules si nécessaire pour un rendu clair et professionnel.
Tu ne modifies jamais les chiffres, dates, unités ou profondeurs.
Tu ne touches pas aux tableaux, calculs ou graphiques.
Tu respectes strictement les espacements, la ponctuation et les retours à la ligne du texte d’origine.
Tu réponds uniquement par le texte corrigé.
"""

# ======================
# FONCTIONS
# ======================
def corriger_bloc(textes):
    try:
        contenu = "\n\n".join(textes)

        response = client.chat.completions.create(
            model="gpt-5-mini",
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": contenu}
            ],
            timeout=60
        )

        return response.choices[0].message.content.split("\n\n")

    except Exception:
        return textes


# ======================
# INTERFACE STREAMLIT
# ======================
st.set_page_config(page_title="Relecture BE", layout="centered")
st.title("🛠️ Relecture & reformulation – EGSOL")

uploaded_file = st.file_uploader(
    "Dépose ton rapport Word (.docx)",
    type=["docx"]
)

if uploaded_file and st.button("Corriger le rapport"):
    with st.spinner("⏳ Correction en cours..."):
        doc = Document(uploaded_file)
        paragraphs = doc.paragraphs[:200]

        bloc_textes = []
        bloc_indices = []

        progress = st.progress(0)
        total = len(paragraphs)
        traites = 0

        for i, paragraph in enumerate(paragraphs):

            # ❌ Ne pas toucher aux paragraphes justifiés
            if paragraph.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                traites += 1
                progress.progress(traites / total)
                continue

            # ❌ Ne pas toucher aux listes Word (puces, tirets, retraits)
            if paragraph._p.pPr is not None and paragraph._p.pPr.numPr is not None:
                traites += 1
                progress.progress(traites / total)
                continue

            texte = paragraph.text.strip()

            if len(texte) > 30:
                bloc_textes.append(texte)
                bloc_indices.append(i)

            if len(bloc_textes) == 10:
                corrections = corriger_bloc(bloc_textes)
                for idx, corr in zip(bloc_indices, corrections):
                    paragraphs[idx].text = corr
                bloc_textes = []
                bloc_indices = []

            traites += 1
            progress.progress(traites / total)

        if bloc_textes:
            corrections = corriger_bloc(bloc_textes)
            for idx, corr in zip(bloc_indices, corrections):
                paragraphs[idx].text = corr

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
        doc.save(tmp.name)

        with open(tmp.name, "rb") as f:
            st.success("✅ Rapport corrigé avec succès")
            st.download_button(
                "📄 Télécharger le rapport corrigé",
                f,
                file_name="rapport_corrige.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
