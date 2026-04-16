import os, re, json, io
from flask import Flask, request, send_file, render_template, jsonify
from difflib import SequenceMatcher
import fitz
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE_DIR, 'catalogue_cashmag.json'), 'r', encoding='utf-8') as f:
    CATALOGUE = json.load(f)

def normalize(s):
    s = (s or '').lower()
    for a, b in [('à','a'),('â','a'),('é','e'),('è','e'),('ê','e'),('î','i'),
                 ('ô','o'),('û','u'),('ù','u'),('ç','c'),("'",' '),('-',' ')]:
        s = s.replace(a, b)
    return re.sub(r'\s+', ' ', s).strip()

def find_best(produit, nic):
    qn = normalize(produit)
    qn = re.sub(r'^one taste |^pack\s+', '', qn)
    qn = re.sub(r'\s*par\s+\d+\b', '', qn).strip()
    nic_str = str(int(nic)) + 'mg' if nic and nic not in ('0','00') else ''
    best, best_score = None, -1
    for p in CATALOGUE:
        lib = normalize(p['libelle'])
        s = SequenceMatcher(None, qn, lib).ratio()
        if nic_str:
            if nic_str in lib: s += 0.25
            else: s -= 0.15
        else:
            if re.search(r'\d+mg', lib) and '0mg' not in lib: s -= 0.1
        words = [w for w in qn.split() if len(w) > 3]
        s += sum(0.05 for w in words if w in lib)
        if s > best_score:
            best_score, best = s, p
    return best, round(best_score, 2)

def parse_bl(text):
    items, seen = [], set()
    lines = text.split('\n')
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(r'(#REF\d+-\d+)(.*)', line)
        if m:
            ref = m.group(1)
            desc_parts = [m.group(2).strip()]
            j = i + 1
            while j < len(lines):
                nl = lines[j].strip()
                if re.match(r'#REF\d+-\d+', nl) or re.match(r'^Page\s*:', nl) or 'Colisage' in nl:
                    break
                if re.match(r'^\d{1,3}$', nl) and j > i + 1:
                    break
                desc_parts.append(nl)
                j += 1
            block = ' '.join(desc_parts)
            qty_m = re.search(r'\)\s*(\d{1,3})\s*$', block) or re.search(r'\)\s+(\d{1,3})', block)
            qty = int(qty_m.group(1)) if qty_m else 0
            if not qty and j < len(lines) and re.match(r'^\d{1,3}$', lines[j].strip()):
                qty = int(lines[j].strip()); j += 1
            nic_m = re.search(r'[Dd]osage\s+[Nn]icotine\s*:\s*(\d+)\s*mg', block)
            nic = str(int(nic_m.group(1))) if nic_m else '0'
            produit = re.sub(r'\s*\(.*', '', block).strip()
            produit = re.sub(r'\s*-\s*(0mg|\d+mg)\s*$', '', produit, flags=re.I).strip()
            key = ref + '|' + nic
            if qty > 0 and key not in seen:
                seen.add(key)
                items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
            i = j
        else:
            i += 1
    return items

def parse_facture(text):
    items = []
    for m in re.finditer(
        r'([A-Z0-9]{6,})\s+(.+?)(?:Nicotine\s*:\s*(\d+)\s*mg[^)]*\)?\s*)?(?:0\s*%|20\s*%)\s+[\d,]+\s*€\s+(\d+)',
        text, re.I):
        ref = m.group(1)
        produit = re.sub(r'\([^)]+\)', '', m.group(2)).strip()
        produit = re.sub(r'\s*-\s*(Pulp\s*-\s*FRC|FR)\s*', ' ', produit, flags=re.I).strip()
        nic = m.group(3) or '0'
        qty = int(m.group(4))
        if qty > 0 and len(ref) >= 6:
            items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
    return items

def parse_doc(text):
    if '#REF' in text and ('Colisage' in text or 'Dosage Nicotine' in text):
        return parse_bl(text)
    return parse_facture(text)

def generate_excel(results, filename):
    VERT = "C6EFCE"; VERT_TXT = "276221"
    ORANGE = "FFEB9C"; ORANGE_TXT = "9C5700"
    BLEU = "1F4E79"; BLEU2 = "2E75B6"
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Entrée stock"
    ws.merge_cells('A1:H1')
    ws['A1'] = f"ENTRÉE DE STOCK — {filename}"
    ws['A1'].font = Font(bold=True, size=12, color='FFFFFF', name='Arial')
    ws['A1'].fill = PatternFill('solid', start_color=BLEU)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 22
    ok_c = sum(1 for r in results if r['statut'] == 'OK')
    ws.merge_cells('A2:H2')
    ws['A2'] = f"{len(results)} références  |  ✓ {ok_c} OK  |  ⚠ {len(results)-ok_c} à vérifier"
    ws['A2'].font = Font(size=10, color='595959', name='Arial')
    ws['A2'].fill = PatternFill('solid', start_color='DEEAF1')
    ws['A2'].alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 15
    for col, h in enumerate(['#','Réf. Fournisseur','Produit Fournisseur','Nicotine','Qté','Libellé Cash Mag','ID Cash Mag','Statut'], 1):
        c = ws.cell(row=3, column=col, value=h)
        c.font = Font(bold=True, color='FFFFFF', name='Arial', size=10)
        c.fill = PatternFill('solid', start_color=BLEU2)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = border
    ws.row_dimensions[3].height = 18
    for i, r in enumerate(results, 1):
        rn = i + 3
        is_av = r['statut'] == 'A VERIFIER'
        nic_label = (r['nic'] + 'mg') if r['nic'] not in ('0','00') else '0mg'
        for col, val in enumerate([i, r['ref'], r['produit'], nic_label, r['qty'],
                                    r['cashMagLibelle'], r['cashMagId'], r['statut']], 1):
            c = ws.cell(row=rn, column=col, value=val)
            c.font = Font(name='Arial', size=10)
            c.border = border
            c.alignment = Alignment(vertical='center', horizontal='center' if col in [1,4,5,7,8] else 'left')
            if col == 8:
                c.fill = PatternFill('solid', start_color=ORANGE if is_av else VERT)
                c.font = Font(name='Arial', size=10, bold=True, color=ORANGE_TXT if is_av else VERT_TXT)
        ws.row_dimensions[rn].height = 16
    for col, w in zip('ABCDEFGH', [4, 18, 34, 10, 6, 40, 13, 13]):
        ws.column_dimensions[col].width = w
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.route('/')
def index():
    return render_template('index.html', nb_produits=len(CATALOGUE))

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400
    file = request.files['file']
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Envoyez un fichier PDF'}), 400
    try:
        doc = fitz.open(stream=file.read(), filetype='pdf')
        text = "\n".join(page.get_text() for page in doc)
        items = parse_doc(text)
        if not items:
            return jsonify({'error': 'Aucune référence trouvée dans ce PDF'}), 400
        results = []
        for item in items:
            match, score = find_best(item['produit'], item['nic'])
            results.append({**item,
                'cashMagLibelle': match['libelle'] if match else 'NON TROUVÉ',
                'cashMagId': str(match['id']) if match else '',
                'score': score,
                'statut': 'OK' if score >= 0.60 else 'A VERIFIER'})
        fname = os.path.splitext(file.filename)[0]
        excel = generate_excel(results, fname)
        return send_file(excel, as_attachment=True,
            download_name=f'entree_stock_{fname}.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
