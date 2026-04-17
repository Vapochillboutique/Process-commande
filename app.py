import streamlit as st
import pandas as pd
import json
import fitz  # PyMuPDF
from difflib import SequenceMatcher
import io
import re
import time

# configuration de la page pour le fond noir et l'interface centrée
st.set_page_config(
    page_title="Vapochill - Correspondance Facture ➔ Cash Mag",
    layout="centered", # Zone de téléchargement au milieu
    initial_sidebar_state="collapsed", # Masque la barre latérale
)

# Thème sombre de Streamlit
st.markdown("""
    <style>
        .stApp { background-color: #0E1117; color: white; }
        .stFileUploader { background-color: #262730; border-radius: 10px; padding: 20px; }
        .stDataFrame { background-color: #1E1E1E; }
        h1 { text-align: center; color: white; }
    </style>
""", unsafe_allow_value=True)

# Affichage du logo
try:
    st.image('Logo Tours.png', width=200, use_container_width=False, output_format="PNG")
except:
    st.error("Le fichier 'Logo Tours.png' est introuvable sur GitHub.")

st.title("📦 Correspondance Facture ➔ Cash Mag")

# --- Logique d'analyse renforcée (Conservée de la version précédente) ---
@st.cache_data
def load_catalogue():
    try:
        with open('catalogue_cashmag.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Erreur catalogue : {e}")
        return []

def normalize(s):
    if not s: return ""
    return "".join(e for e in str(s).lower() if e.isalnum())

def extract_specs(text):
    res = {"mg": None, "ml": None}
    text_lower = text.lower()
    mg_match = re.search(r'(\d{1,2})\s?mg', text_lower)
    if mg_match: res["mg"] = mg_match.group(1).zfill(2)
    ml_match = re.search(r'(\d{2,3})\s?ml', text_lower)
    if ml_match: res["ml"] = ml_match.group(1)
    return res

def find_best_match(text, catalogue):
    specs_facture = extract_specs(text)
    text_norm = normalize(text)
    best_match, highest_score = None, 0
    
    for item in catalogue:
        libelle_cat = item.get('libelle', '')
        specs_cat = extract_specs(libelle_cat)
        lib_norm = normalize(libelle_cat)
        
        # Filtre de sécurité ml/mg (Critique pour Pulp)
        if specs_facture["ml"] and specs_cat["ml"] and specs_facture["ml"] != specs_cat["ml"]: continue
        if specs_facture["mg"] and specs_cat["mg"] and specs_facture["mg"] != specs_cat["mg"]: continue
        
        score = SequenceMatcher(None, text_norm, lib_norm).ratio()
        
        # Bonus sur les mots clés du parfum
        if len(set(text_norm.split()).intersection(set(lib_norm.split()))) > 1: score += 0.1
            
        if score > highest_score:
            highest_score, best_match = score, item
            
    return best_match, highest_score

# --- Interface Utilisateur (Version Sobres) ---
catalogue = load_catalogue()

# Zone de téléchargement centrée
uploaded_file = st.file_uploader("📂 Dépose ta facture PDF ici", type="pdf")

if uploaded_file and catalogue:
    # Barre de chargement animée
    with st.spinner('Analyse technique des produits en cours...'):
        my_bar = st.progress(0)
        
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        total_lines = sum(len(page.get_text("text").split('\n')) for page in doc)
        processed_lines = 0
        extracted_data = []
        
        for page in doc:
            lines = page.get_text("text").split('\n')
            for line in lines:
                processed_lines += 1
                line = line.strip()
                # On filtre les lignes qui ressemblent à des produits
                if len(line) > 8 and not any(x in line.lower() for x in ['total', 'facture', 'iban', 'tva', 'client', 'page']):
                    match, score = find_best_match(line, catalogue)
                    if match and score > 0.45: # Seuil de confiance relevé pour plus de précision
                        extracted_data.append({
                            "Ligne Facture": line,
                            "PRODUIT CASH MAG": match.get('libelle'),
                            "ID": match.get('id'),
                            "Fiabilité": f"{int(score*100)}%"
                        })
                
                # Mise à jour de la barre de progression
                if processed_lines % 5 == 0:
                    percent_complete = int(processed_lines / total_lines * 100)
                    my_bar.progress(min(percent_complete, 100))
        
        time.sleep(0.5) # Pour l'effet visuel de fin de chargement

        # Affichage du rendu direct sur le site
        if extracted_data:
            df = pd.DataFrame(extracted_data)
            st.success("Analyse terminée. Voici les correspondances :")
            
            # Tableau direct sur le site
            st.dataframe(df, use_container_width=True)
            
            # Bouton de téléchargement
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Télécharger le fichier Excel", output.getvalue(), "import_cashmag.xlsx")
        else:
            st.warning("Aucun produit reconnu. Vérifie que le PDF est bien une facture.")
            
elif not catalogue:
    st.error("Le catalogue n'a pas pu être chargé. Vérifie le fichier .json sur GitHub.")
