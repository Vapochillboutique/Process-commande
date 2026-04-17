import streamlit as st
import pandas as pd
import fitz
from difflib import SequenceMatcher
import io
import re
import os
import time

# Config de la page pour le fond noir
st.set_page_config(page_title="Vapochill Matching", layout="centered")

st.markdown("""
    <style>
    .stApp { background-color: #0E1117; color: white; }
    .stDataFrame { background-color: #1E1E1E; }
    div.stButton > button { width: 100%; border-radius: 5px; }
    </style>
    """, unsafe_allow_html=True)

# Logo centré
if os.path.exists('Logo Tours.png'):
    st.image('Logo Tours.png', width=200)
else:
    st.title("VAPOCHILL TOURS")

@st.cache_data
def load_data():
    if not os.path.exists('catalogue.csv'):
        return None
    try:
        # On lit le fichier en sautant les lignes vides du début (environ 48)
        # On cherche la ligne d'en-tête contenant "Libellé"
        raw_df = pd.read_csv('catalogue.csv', sep=None, engine='python', header=None, on_bad_lines='skip')
        header_idx = raw_df[raw_df.apply(lambda r: r.astype(str).str.contains('Libellé').any(), axis=1)].index[0]
        
        df = pd.read_csv('catalogue.csv', sep=None, engine='python', skiprows=header_idx)
        # Nettoyage : Retire la FIDELITE et les symboles $
        df = df[~df['Rubrique'].str.contains('FIDÉLITÉ', na=False, case=False)]
        df['Libellé'] = df['Libellé'].str.replace('$', '', regex=False)
        return df[['Libellé', '#ID', 'Rubrique']]
    except Exception as e:
        st.error(f"Erreur catalogue : {e}")
        return None

def extract_specs(text):
    t = text.lower()
    mg = re.search(r'(\d{1,2})\s?mg', t)
    ml = re.search(r'(\d{2,3})\s?ml', t)
    return {"mg": mg.group(1).zfill(2) if mg else None, "ml": ml.group(1) if ml else None}

def find_match(text, df_cat):
    spec_f = extract_specs(text)
    text_norm = "".join(e for e in text.lower() if e.isalnum())
    
    # Pré-filtrage par contenance/nicotine pour la rapidité
    temp = df_cat.copy()
    if spec_f["ml"]:
        temp = temp[temp['Libellé'].str.contains(spec_f["ml"] + "ml", case=False, na=False)]
    
    best_lib, best_id, max_score = "NON TROUVÉ", "", 0
    for _, row in temp.iterrows():
        cat_norm = "".join(e for e in str(row['Libellé']).lower() if e.isalnum())
        score = SequenceMatcher(None, text_norm, cat_norm).ratio()
        if score > max_score:
            max_score, best_lib, best_id = score, row['Libellé'], row['#ID']
            if score > 0.95: break
            
    return best_lib, best_id, max_score

df_cat = load_data()

if df_cat is not None:
    st.write(f"✅ Catalogue prêt : **{len(df_cat)} produits** chargés.")
    uploaded_file = st.file_uploader("📂 Dépose ta facture PDF ici", type="pdf")

    if uploaded_file:
        progress_bar = st.progress(0)
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        results = []
        
        all_pages = list(doc)
        for i, page in enumerate(all_pages):
            lines = page.get_text("text").split('\n')
            for line in lines:
                line = line.strip()
                if len(line) > 10 and not any(x in line.lower() for x in ['total', 'tva', 'iban', 'page']):
                    lib, id_cm, score = find_match(line, df_cat)
                    if score > 0.40:
                        results.append({"Facture": line, "NOM CASH MAG": lib, "ID": id_cm, "Score": f"{int(score*100)}%"})
            progress_bar.progress((i + 1) / len(all_pages))
            
        if results:
            st.success("Analyse terminée !")
            res_df = pd.DataFrame(results).drop_duplicates(subset=['Facture'])
            st.dataframe(res_df, use_container_width=True)
            
            output = io.BytesIO()
            res_df.to_excel(output, index=False)
            st.download_button("📥 Télécharger l'Excel pour Cash Mag", output.getvalue(), "import_stock.xlsx")
else:
    st.warning("⚠️ En attente du fichier 'catalogue.csv' sur GitHub...")
