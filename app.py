import streamlit as st
import pandas as pd
import json
import fitz  # PyMuPDF
from difflib import SequenceMatcher
import io

st.set_page_config(page_title="Vapochill - Correspondance Cash Mag", layout="wide")

st.title("📦 Vapochill : Correspondance Facture ➔ Cash Mag")

# Chargement du catalogue
@st.cache_data
def load_catalogue():
    try:
        with open('catalogue_cashmag.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return []

catalogue = load_catalogue()

def normalize(s):
    if not s: return ""
    return "".join(e for e in str(s).lower() if e.isalnum())

def find_best_match(text, catalogue):
    text_norm = normalize(text)
    best_match = None
    highest_score = 0
    
    for item in catalogue:
        lib_norm = normalize(item.get('libelle', ''))
        score = SequenceMatcher(None, text_norm, lib_norm).ratio()
        if score > highest_score:
            highest_score = score
            best_match = item
    return best_match, highest_score

# Interface de téléchargement
uploaded_file = st.file_uploader("Glisse ta facture PDF ici", type="pdf")

if uploaded_file and catalogue:
    with st.spinner('Analyse de la facture en cours...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        extracted_data = []
        
        for page in doc:
            lines = page.get_text("text").split('\n')
            for line in lines:
                line = line.strip()
                # On filtre les lignes qui ressemblent à des produits (plus de 10 caractères)
                if len(line) > 10 and not any(x in line.lower() for x in ['total', 'facture', 'iban', 'tva']):
                    match, score = find_best_match(line, catalogue)
                    if match and score > 0.5:
                        extracted_data.append({
                            "Produit Facture": line,
                            "NOM À COPIER (CASH MAG)": match.get('libelle'),
                            "ID": match.get('id'),
                            "Fiabilité": f"{int(score*100)}%"
                        })

        if extracted_data:
            df = pd.DataFrame(extracted_data)
            st.success(f"{len(df)} produits identifiés !")
            st.table(df) # Affiche le tableau directement
            
            # Bouton pour télécharger en Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Télécharger le tableau Excel", output.getvalue(), "correspondance.xlsx")
        else:
            st.warning("Aucun produit reconnu. Vérifie que le PDF est bien une facture.")
elif not catalogue:
    st.error("Le fichier catalogue_cashmag.json est vide ou introuvable sur GitHub.")
