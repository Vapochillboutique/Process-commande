import streamlit as st
import pandas as pd
import json
import fitz
from difflib import SequenceMatcher
import io
import re
import time
import os

st.set_page_config(page_title="Vapochill - Matching", layout="centered")

# Style pour le fond noir
st.markdown("""
    <style>
    .stApp { background-color: #0E1117; color: white; }
    h1 { text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# Affichage Logo
if os.path.exists('Logo Tours.png'):
    st.image('Logo Tours.png', width=200)
else:
    st.title("VAPOCHILL")

st.write("---")

@st.cache_data
def load_catalogue():
    try:
        with open('catalogue_cashmag.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except: return []

def extract_specs(text):
    res = {"mg": None, "ml": None}
    t = text.lower()
    mg = re.search(r'(\d{1,2})\s?mg', t)
    if mg: res["mg"] = mg.group(1).zfill(2)
    ml = re.search(r'(\d{2,3})\s?ml', t)
    if ml: res["ml"] = ml.group(1)
    return res

def find_match(text, catalogue):
    spec_f = extract_specs(text)
    best, score = None, 0
    for item in catalogue:
        spec_c = extract_specs(item['libelle'])
        if spec_f["ml"] and spec_c["ml"] and spec_f["ml"] != spec_c["ml"]: continue
        if spec_f["mg"] and spec_c["mg"] and spec_f["mg"] != spec_c["mg"]: continue
        
        s = SequenceMatcher(None, text.lower(), item['libelle'].lower()).ratio()
        if s > score: score, best = s, item
    return best, score

catalogue = load_catalogue()

# Zone de téléchargement au milieu
uploaded_file = st.file_uploader("📂 Dépose ta facture PDF ici", type="pdf")

if uploaded_file and catalogue:
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    results = []
    
    pages = list(doc)
    for i, page in enumerate(pages):
        lines = page.get_text("text").split('\n')
        for line in lines:
            if len(line.strip()) > 10:
                m, s = find_match(line, catalogue)
                if m and s > 0.45:
                    results.append({"Facture": line, "CASH MAG": m['libelle'], "ID": m['id']})
        
        # Mise à jour barre de chargement
        progress = (i + 1) / len(pages)
        progress_bar.progress(progress)
        status_text.text(f"Analyse de la page {i+1}/{len(pages)}...")

    if results:
        st.success("Analyse terminée !")
        df = pd.DataFrame(results)
        st.table(df) # Affichage direct sur le site
        
        output = io.BytesIO()
        df.to_excel(output, index=False)
        st.download_button("📥 Télécharger l'Excel", output.getvalue(), "import_cashmag.xlsx")
