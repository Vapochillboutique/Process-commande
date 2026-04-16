import os, re, json, io
from flask import Flask, request, send_file
from difflib import SequenceMatcher
import fitz  # PyMuPDF
import openpyxl

app = Flask(__name__)

# Chargement du catalogue
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
catalogue_path = os.path.join(BASE_DIR, 'catalogue_cashmag.json')
CATALOGUE = []
if os.path.exists(catalogue_path):
    with open(catalogue_path, 'r', encoding='utf-8') as f:
        CATALOGUE = json.load(f)

def normalize(s):
    if not s: return ""
    return re.sub(r'[^a-z0-9]', '', str(s).lower())

def find_best(name):
    norm_name = normalize(name)
    best, best_score = None, 0
    for p in CATALOGUE:
        lib = normalize(p.get('libelle', ''))
        score = SequenceMatcher(None, norm_name, lib).ratio()
        if score > best_score:
            best_score, best = score, p
    return best, best_score

@app.route('/')
def index():
    return f"Serveur pret. Catalogue: {len(CATALOGUE)} produits."

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files: return "Pas de fichier", 400
    file = request.files['file']
    doc = fitz.open(stream=file.read(), filetype="pdf")
    results = []
    for page in doc:
        for line in page.get_text("text").split('\n'):
            if len(line.strip()) > 10:
                match, score = find_best(line)
                if match and score > 0.5:
                    results.append([line, match['libelle'], match['id']])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Nom Facture', 'NOM CASH MAG', 'ID CASH MAG'])
    for r in results: ws.append(r)
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name='correspondance.xlsx')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
