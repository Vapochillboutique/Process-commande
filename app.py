import os, re, json, io
from flask import Flask, request, send_file, render_template, jsonify
from difflib import SequenceMatcher
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Chargement sécurisé du catalogue
CATALOGUE = []
catalogue_path = os.path.join(BASE_DIR, 'catalogue_cashmag.json')
if os.path.exists(catalogue_path):
    with open(catalogue_path, 'r', encoding='utf-8') as f:
        CATALOGUE = json.load(f)

def normalize(s):
    if not s: return ""
    s = s.lower()
    s = re.sub(r'[éèêë]', 'e', s)
    s = re.sub(r'[àâä]', 'a', s)
    s = re.sub(r'[îï]', 'i', s)
    s = re.sub(r'[ôö]', 'o', s)
    s = re.sub(r'[ûüù]', 'u', s)
    s = re.sub(r'[^a-z0-9\s]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()

def find_best(produit_facture):
    name_norm = normalize(produit_facture)
    # Détection nicotine (ex: 3mg, 6mg, 12mg)
    nic_match = re.search(r'(\d{1,2})\s*mg', name_norm)
    nic_val = nic_match.group(1).zfill(2) + "mg" if nic_match else None

    best, best_score = None, -1
    for p in CATALOGUE:
        lib_cm = p.get('libelle', '')
        lib_norm = normalize(lib_cm)
        score = SequenceMatcher(None, name_norm, lib_norm).ratio()
        
        # Bonus nicotine
        if nic_val:
            if nic_val in lib_norm or nic_val.replace('0', '') in lib_norm:
                score += 0.25
            elif "mg" in lib_norm:
                score -= 0.20
        
        if score > best_score:
            best_score, best = score, p
    return best, round(best_score, 2)

def generate_excel(results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Correspondance Cash Mag"
    headers = ['Produit Facture', 'NOM À COPIER CASH MAG', 'ID Cash Mag', 'Statut']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F4E79")

    for i, r in enumerate(results, 2):
        ws.cell(row=i, column=1, value=r['produit'])
        ws.cell(row=i, column=2, value=r['cashMagLibelle'])
        ws.cell(row=i, column=3, value=r['cashMagId'])
        status_cell = ws.cell(row=i, column=4, value=r['statut'])
        if r['statut'] == 'OK':
            status_cell.fill = PatternFill("solid", start_color="C6EFCE")
        else:
            status_cell.fill = PatternFill("solid", start_color="FFEB9C")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/')
def index():
    return f"Application active. Catalogue chargé : {len(CATALOGUE)} produits."

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "Aucun fichier", 400
    
    file = request.files['file']
    doc = fitz.open(stream=file.read(), filetype="pdf")
    all_lines = []
    for page in doc:
        all_lines.extend(page.get_text("text").split('\n'))
    
    results = []
    for line in all_lines:
        line = line.strip()
        # On ignore les lignes trop courtes ou les mentions inutiles
        if len(line) < 12 or any(x in line.lower() for x in ['total', 'facture', 'pge', 'iban', 'tva']):
            continue
            
        match, score = find_best(line)
        if match and score > 0.40: # On ne garde que ce qui ressemble à un produit
            results.append({
                'produit': line,
                'cashMagLibelle': match['libelle'],
                'cashMagId': match.get('id', ''),
                'statut': 'OK' if score > 0.70 else 'A VERIFIER'
            })
    
    excel = generate_excel(results)
    return send_file(excel, as_attachment=True, download_name='correspondance_stock.xlsx')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
