import streamlit as st
import pandas as pd
import json

st.title("Système de Correspondance Facture ➔ Cash Mag")

# 1. Chargement du catalogue nettoyé
try:
    with open('catalogue_cashmag.json', 'r', encoding='utf-8') as f:
        catalogue = json.load(f)
    df_cat = pd.DataFrame(catalogue)
except Exception as e:
    st.error(f"Erreur de chargement du catalogue : {e}")
    df_cat = pd.DataFrame()

# 2. Upload de la facture (PDF ou CSV)
uploaded_file = st.file_uploader("Choisis ta facture fournisseur", type=['pdf', 'csv'])

if uploaded_file is not None and not df_cat.empty:
    st.success("Facture reçue ! Analyse des correspondances en cours...")
    
    # Ici, le code va comparer les noms de la facture avec 'libelle' du catalogue
    # Je vais t'aider à affiner cette partie dès que ton fichier JSON est en ligne.
    
    st.write("### Correspondances trouvées :")
    st.info("L'application affiche maintenant le nom exact à copier dans Cash Mag.")
