import streamlit as st
import pandas as pd
import json
import fitz  # PyMuPDF
from difflib import SequenceMatcher
import io
import re

st.set_page_config(page_title="Vapochill - Correspondance Précise", layout="wide")

st.title("📦 Vapochill : Correspondance Facture ➔ Cash Mag")
st.subheader("Analyse sécurisée (Contenance & Nicotine)")

# Chargement du catalogue
@st.cache_data
def load_catalogue():
    try:
        with open('catalogue_cashmag.json', 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        st.error(f"Erreur de lecture du catalogue : {e}")
        return []

def normalize(s):
    if not s: return ""
    # On garde les chiffres et lettres pour ne pas perdre les "10ml" ou "6mg"
    return "".join(e for e in str(s).lower() if e.isalnum())

def extract_specs(text):
    """Extrait précisément le taux de nicotine et la contenance"""
    res = {"mg": None, "ml": None}
    text_lower = text.lower()
    
    # Extraction MG (nicotine) : cherche "3mg", "06mg", "12 mg"
    mg_match = re.search(r'(\d{1,2})\s?mg', text_lower)
    if mg_match: 
        res["mg"] = mg_match.group(1).zfill(2) # "06" au lieu de "6" pour comparaison stricte
    
    # Extraction ML (contenance) : cherche "10ml", "60 ml", "200ml"
    ml_match = re.search(r'(\d{2,3})\s?ml', text_lower)
    if ml_match: 
        res["ml"] = ml_match.group(1)
        
    return res

def find_best_match(text, catalogue):
    specs_facture = extract_specs(text)
    text_norm = normalize(text)
    
    best_match = None
    highest_score = 0
    
    for item in catalogue:
        libelle_cat = item.get('libelle', '')
        specs_cat = extract_specs(libelle_cat)
        lib_norm = normalize(libelle_cat)
        
        # --- FILTRE DE SÉCURITÉ ---
        # Si on a trouvé des ML des deux côtés et qu'ils sont différents -> ON REJETTE
        if specs_facture["ml"] and specs_cat["ml"] and specs_facture["ml"] != specs_cat["ml"]:
            continue
            
        # Si on a trouvé des MG des deux côtés et qu'ils sont différents -> ON REJETTE
        if specs_facture["mg"] and specs_cat["mg"] and specs_facture["mg"] != specs_cat["mg"]:
            continue
        
        # Calcul de similarité
        score = SequenceMatcher(None, text_norm, lib_norm).ratio()
        
        # Bonus si des mots clés importants (comme le parfum) sont présents
        # Cela aide à différencier "Fruit rouge" de "Christmas Cookie"
        words_facture = set(text_norm.split())
        words_cat = set(lib_norm.split())
        common_words = words_facture.intersection(words_cat)
        score += (len(common_words) * 0.05)

        if score > highest_score:
            highest_score = score
            best_match = item
            
    return best_match, highest_score

# Interface
catalogue = load_catalogue()
uploaded_file = st.file_uploader("Glisse ta facture PDF ici", type="pdf")

if uploaded_file and catalogue:
    with st.spinner('Analyse technique des produits en cours...'):
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        extracted_data = []
        
        for page in doc:
            lines = page.get_text("text").split('\n')
            for line in lines:
                line = line.strip()
                # On ignore les lignes trop courtes ou contenant des mots financiers
                if len(line) > 8 and not any(x in line.lower() for x in ['total', 'facture', 'iban', 'tva', 'client', 'page']):
                    match, score = find_best_match(line, catalogue)
                    if match and score > 0.4: # Seuil de confiance
                        extracted_data.append({
                            "Ligne Facture": line,
                            "PRODUIT CASH MAG": match.get('libelle'),
                            "ID": match.get('id'),
                            "Confiance": f"{int(score*100)}%"
                        })

        if extracted_data:
            df = pd.DataFrame(extracted_data)
            st.success("Analyse terminée. Vérifie les correspondances ci-dessous :")
            st.dataframe(df, use_container_width=True)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            st.download_button("📥 Télécharger pour Cash Mag (Excel)", output.getvalue(), "import_cashmag.xlsx")
        else:
            st.warning("Aucun produit n'a pu être associé. Vérifie ton catalogue ou le PDF.")
