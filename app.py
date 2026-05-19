import os, re, json, io, uuid
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
    s = s.replace('\u00e0','a').replace('\u00e2','a').replace('\u00e9','e').replace('\u00e8','e')
    s = s.replace('\u00ea','e').replace('\u00ee','i').replace('\u00f4','o').replace('\u00fb','u')
    s = s.replace('\u00f9','u').replace('\u00e7','c').replace("'",' ').replace('-',' ')
    return re.sub(r'\s+', ' ', s).strip()

def extract_volume(s):
    m = re.search(r'(\d+)\s*ml', (s or '').lower())
    return int(m.group(1)) if m else None

def extract_ohm(s):
    s_low = s.lower().replace(',','.')
    m = re.search(r'(?:valeur|ohm)\s*:?\s*(?:dm\d+\s+)?([01]\.[0-9]{1,2})\s*(?:ohm|\?)', s_low)
    if m: return m.group(1)
    matches = re.findall(r'([01]\.[0-9]{1,2})\s*ohm', s_low)
    if matches: return matches[-1]
    matches = re.findall(r'([01]\.[0-9]{1,2})\?', s_low)
    if matches: return matches[-1]
    return None

BRUIT = {'ultimate','sweet','edition','green','zero','classic','tabac','gourmand',
         'by','maison','le','pod','liquide','fizz','concentre','arome',
         'aromes','liquides','et','du','de','la','les','par','top','fill','unik'}

def get_name_words(s):
    s = normalize(s)
    s = re.sub(r'^pack\s+|^concentre\s+|^e liquide\s+', '', s)
    s = re.sub(r'\s*\d+\s*ml.*', '', s)
    for marque in ['aromes et liquides','a&l','fighter fuel','pulp','cupide','savourea',
                   'maison fuel','enfer','aspire','vaporesso','voopoo','geekvape','eleaf',
                   'le french liquide','lemon time','fruizee','vampire vape','swoke',
                   'salt e vapor','liquideo','tjuice','lost vape','lostvape','polaris',
                   'tribal force','jnr','multi freeze','greenvillage','justfog','smoketech',
                   'eliquid france','gfc provap','adns','vape cellar','l absolv',
                   'white rabbit','avap','wpuff','le petit verger','petit verger','lpv']:
        s = s.replace(normalize(marque), '')
    return [w for w in s.split() if len(w) > 2 and w not in BRUIT]

INDEX = {}
for i, p in enumerate(CATALOGUE):
    lib = normalize(p['libelle'])
    for w in set(lib.split()):
        if len(w) > 2:
            if w not in INDEX: INDEX[w] = []
            INDEX[w].append(i)

print(f"Catalogue: {len(CATALOGUE)} produits | Index: {len(INDEX)} mots")

# ── Correspondances manuelles ──────────────────────────────────────────────────
MANUELS = {
    normalize('Lemon 50ml Lemon'): 'lemon time lemon',
    normalize('Lemon time Lemon 50ml'): 'lemon time lemon',
    normalize('Tireboulette'): 'tireboulette',
    normalize('Tireboulette Peche Mangue Passion'): 'tireboulette',
    normalize('Corossol Peche 50ml Le Petit Verger'): 'corossol peche lpv 50ml',
    normalize('Corossol Peche 50ml Savourea'): 'corossol peche lpv 50ml',
    normalize('Corossol Peche 30ml'): 'corossol peche lpv 30ml',
    normalize('Concentre Corossol Peche 30ml'): 'corossol peche lpv 30ml',
    normalize('Cartouches vides Veynom Air'): 'cartouche vide veynom aspire',
    normalize('Pyrex Centaurus Sub Ohm Bubble'): 'pyrex centaurus bubble',
    normalize('Pyrex Centaurus Sub Ohm'): 'pyrex centaurus bubble',
    normalize('Centaurus Sub Ohm Bubble'): 'pyrex centaurus bubble',
    normalize('Clearomiseur Centaurus Sub Ohm V2'): 'pyrex centaurus bubble',
}

# ── Table résistances/cartouches ──────────────────────────────────────────────
RESISTANCE_MAP = {
    'cartouches apex': 'cartouches apex',
    'cartouche apex': 'cartouches apex',
    'b series': 'resistance nano',
    'z xm': 'resistance zeus sub',
    'zenith': 'resistance z coil',
    'z coil': 'resistance z coil',
    'nautilus': 'resistance nautilus',
    'pnp x': 'resistance pnp',
    'pnp': 'resistance pnp',
    'luxe xr': 'cartouche luxe xr',
    'luxe x': 'cartouche luxe x',
    'pixo': 'cartouche pixo',
    'soul': 'cartouche geekvape soul',
    'ursa nano': 'cartouche ursa nano',
    'ursa v3': 'cartouche ursa nano',
    'ursa v2': 'cartouche ursa nano',
    'ub max': 'resistances ub max',
    'gtx': 'resistance gtx',
    'veynom air': 'cartouche vide veynom aspire',
    'tfv18': 'resistance tfv18',
    'tpp': 'resistance tpp',
}

def find_resistance_match(produit):
    prod_low = produit.lower()
    ohm = extract_ohm(produit)
    for keyword, cm_prefix in RESISTANCE_MAP.items():
        if keyword in prod_low:
            candidates = []
            for p in CATALOGUE:
                lib = p['libelle'].lower().replace(',','.')
                if cm_prefix in lib:
                    if ohm:
                        if ohm in lib:
                            candidates.append(p)
                    else:
                        candidates.append(p)
            if candidates:
                return candidates[0], 0.99, ohm
    return None, 0, None

# ── Table Airmust/Paperland/Le Primeur ────────────────────────────────────────
AIRMUST_MAP = {
    'bonbon cola': 'bonbon cola 50 ml unik',
    'caramel fondant': 'caramel fondant 50 ml unik',
    'cassis': 'cassis 50 ml unik',
    'cerise intense': 'cerise intense 50 ml unik',
    'citron givre': 'citron givre 50 ml unik',
    'custard vanille': 'custard vanille 50 ml unik',
    'fruit du dragon': 'fruit du dragon 50 ml unik',
    'fruits rouges': 'fruits rouges 50 ml unik',
    'mangue': 'mangue 50 ml unik',
    'menthe du jardin': 'menthe du jardin 50 ml unik',
    'menthe glaciale': 'menthe glaciale 50 ml unik',
    'noisette': 'noisette 50 ml unik',
    'peche': 'peche 50 ml unik',
    'poire': 'poire 50 ml unik',
    'pomme harmonie': 'pomme harmonie 50 ml unik',
    'pop corn': 'popcorn 50 ml unik',
    'popcorn': 'popcorn 50 ml unik',
    'pure passion': 'pure passion 50 ml unik',
    'raisin noir': 'raisin noir 50 ml unik',
    'fraise sauvage': 'fraise 50 ml unik',
    'aspik': 'ferox aspik airmust',
    'hippox': 'ferox hippox airmust',
    'berry pulse': 'berry pulse 50ml paperland',
    'burning blue': 'burning blue 50 ml paperland',
    'golden bless': 'golden bless 50ml paperland',
    'green fizz': 'green fizz 50 ml paperland',
    'navy drop': 'navy drop 50 ml paperland',
    'peach idyll': 'peach idyll 50ml paperland',
    'red lover': 'red lover 50 ml paperland',
    'ruby crush': 'ruby crush 50ml paperland',
    'white dragon': 'white dragon 50 ml paperland',
    'yellow tropic': 'yellow tropic 50 ml paperland',
    'ananas passion': 'ananas passion le primeur 50ml',
    'cassis mangue': 'cassis mangue 50 ml le primeur',
    'cassis pasteque': 'cassis pasteque le primeur 50ml',
    'cerise groseille': 'cerise groseille le primeur 50ml',
    'kiwi banane': 'kiwi banane le primeur 50ml',
    'pitaya framboise': 'pitaya framboise le primeur 50ml',
}

def find_airmust_match(produit):
    prod_n = normalize(produit)
    prod_n = prod_n.replace('unik', 'airmust')
    prod_n = re.sub(r'\s*\d+\s*ml', '', prod_n).strip()
    for keyword, cm_search in AIRMUST_MAP.items():
        if keyword in prod_n:
            cm_words = [w for w in normalize(cm_search).split() if len(w) > 2]
            best, best_score = None, 0
            for p in CATALOGUE:
                lib = normalize(p['libelle'])
                hits = sum(1 for w in cm_words if w in lib)
                if hits > best_score:
                    best_score, best = hits, p
            if best and best_score >= 2:
                return best, 0.98
    return None, 0

# ── Table CBD AG Consulting ───────────────────────────────────────────────────
AG_CBD_MAP = {
    'sb amnesia b20': 'amnesia b20 1gr',
    'amnesia b20': 'amnesia b20 1gr',
    'sb white widow b20': 'white widow sb b20 1gr',
    'white widow b20': 'white widow sb b20 1gr',
    'sb blueberry b20': 'blueberry sb b20 1gr',
    'blueberry b20': 'blueberry b20 1gr',
    'sb cannatonic b20': 'cannatonic sb b20 1gr',
    'cannatonic b20': 'cannatonic sb b20 1gr',
    'sb cannatonic cbd': 'cannatonic sb b20 1gr',
    'sb jack herrer b20': 'jack herrer 1gr',
    'jack herrer b20': 'jack herer b20 1gr',
    'jack herrer': 'jack herrer 1gr',
    'sb gorilla glue b20': 'gorilla glue sb b20 1gr',
    'gorilla glue b20': 'gorilla glue b20 1gr',
    'gorilla glue': 'gorilla glue 1gr',
    'biscotti b20': 'biscotti b20 1gr',
    'biscotti': 'biscotti b20 1gr',
    'papaya sunset b20': 'papaya sunset b20 1gr',
    'papaya sunset': 'papaya sunset b20 1gr',
    'gmo cookies b20': 'gmo cookie b20 1gr',
    'gmo cookie b20': 'gmo cookie b20 1gr',
    'banana jelly b20': 'banana jelly b20 1gr',
    'banana jelly': 'banana jelly 1gr',
    '3x filtre b20': '3x filtre b20 1gr',
    '3x filtr': '3x filtr',
    'sb blueberry cbd': 'blueberry sb b20 1gr',
    'blueberry cbd': 'blueberry sb b20 1gr',
    'bubble hash b52': 'bubble hash b52 1gr',
    'banana capione': 'banana capione 1gr',
    'apple skunk': 'apple skunk 1gr',
    'watermelon thv2': 'watermelon thv2 1gr',
    'gary payton thv2': 'gary payton thv2 1gr',
    'gary payton b20': 'gary payton b20 1gr',
    'watermelon b20': 'watermelon b20 1gr',
}

def find_ag_cbd_match(produit):
    prod_n = normalize(produit).lower()
    prod_n = re.sub(r'lot n[°o]\w+', '', prod_n)
    prod_n = re.sub(r'gtin[\s:]+\d+', '', prod_n)
    prod_n = re.sub(r'sommites florales.*', '', prod_n)
    prod_n = re.sub(r'thc.*', '', prod_n)
    prod_n = re.sub(r'ddm.*', '', prod_n)
    prod_n = prod_n.strip()
    for keyword, cm_search in AG_CBD_MAP.items():
        if keyword in prod_n:
            cm_words = [w for w in normalize(cm_search).split() if len(w) > 1]
            best, best_score = None, 0
            for p in CATALOGUE:
                if 'cannabis' in p['libelle'].lower() or 'ext.' in p['libelle'].lower():
                    lib = normalize(p['libelle'])
                    hits = sum(1 for w in cm_words if w in lib)
                    if hits > best_score:
                        best_score, best = hits, p
            if best and best_score >= 2:
                return best, 0.98
    return None, 0

# ── find_best ─────────────────────────────────────────────────────────────────
def find_best(produit, nic):
    prod_norm = normalize(produit)

    # Résistances et cartouches par OHM
    if ('luxe xr' in prod_norm or 'luxe x' in prod_norm) and 'cartouche' in prod_norm:
        if 'dtl' in prod_norm:
            for p in CATALOGUE:
                if 'luxe xr dtl' in normalize(p['libelle']) and 'cartouche' in normalize(p['libelle']):
                    return p, 0.99
        elif 'mtl' in prod_norm:
            for p in CATALOGUE:
                if 'luxe xr mtl' in normalize(p['libelle']) and 'cartouche' in normalize(p['libelle']):
                    return p, 0.99

    if 'apex' in prod_norm and ('kit' in prod_norm or 'pack' in prod_norm) and 'cartouche' not in prod_norm:
        couleurs = {'black': 'black', 'navy blue': 'navy blue', 'pearl white': 'pearl white',
                    'snow pink': 'snow pink', 'sky blue': 'sky blue', 'silver': 'satin silver',
                    'midnight black': 'black'}
        for coul_fr, coul_cm in couleurs.items():
            if coul_fr in prod_norm:
                for p in CATALOGUE:
                    if 'apex vaporesso' in normalize(p['libelle']) and coul_cm in normalize(p['libelle']):
                        return p, 0.99
        for p in CATALOGUE:
            if 'apex vaporesso' in normalize(p['libelle']):
                return p, 0.95

    res_match, res_score, res_ohm = find_resistance_match(produit)
    if res_match:
        return res_match, res_score

    ag_match, ag_score = find_ag_cbd_match(produit)
    if ag_match:
        return ag_match, ag_score

    airmust_match, airmust_score = find_airmust_match(produit)
    if airmust_match:
        return airmust_match, airmust_score

    for cle, valeur in MANUELS.items():
        if cle in prod_norm:
            for p in CATALOGUE:
                if valeur in normalize(p['libelle']):
                    vol_f = re.search(r'(\d+)\s*ml', produit.lower())
                    vol_c = re.search(r'(\d+)\s*ml', p['libelle'].lower())
                    if vol_f and vol_c and vol_f.group(1) == vol_c.group(1):
                        return p, 0.99
                    elif not vol_f:
                        return p, 0.95
            break

    qn_raw = prod_norm
    qn_raw = re.sub(r'^pack\s+|^concentre\s+', '', qn_raw)
    qn_raw = re.sub(r'\s*par\s+\d+\b', '', qn_raw)
    qn_raw = qn_raw.replace('aromes et liquides','a&l').replace('aromes liquides','a&l')
    qn_raw = qn_raw.replace('le petit verger','lpv').replace('petit verger','lpv').replace('savourea','lpv')
    qn_raw = re.sub(r'\s+', ' ', qn_raw).strip()

    vol_fourn = extract_volume(produit)
    if 'pulp' in qn_raw and vol_fourn == 60:
        qn_raw = qn_raw.replace('60ml','50ml')
        vol_fourn = 50

    nic_str = str(int(nic)) + 'mg' if nic and nic not in ('0','00') else ''
    name_words = get_name_words(produit)
    all_words = [w for w in qn_raw.split() if len(w) > 2]
    candidates = {}
    for w in all_words + name_words:
        for idx in INDEX.get(w, []):
            candidates[idx] = candidates.get(idx, 0) + 1
    if not candidates:
        candidates = {i: 0 for i in range(len(CATALOGUE))}
    top = sorted(candidates.items(), key=lambda x: -x[1])[:300]
    best, best_score = None, -1
    for idx, _ in top:
        p = CATALOGUE[idx]
        lib = normalize(p['libelle'])
        lib_al = lib.replace('aromes et liquides','a&l')
        s = max(SequenceMatcher(None, qn_raw, lib).ratio(),
                SequenceMatcher(None, qn_raw, lib_al).ratio())
        if nic_str:
            if nic_str in lib: s += 0.25
            else: s -= 0.20
        else:
            if re.search(r'\d+mg', lib) and '0mg' not in lib: s -= 0.15
        s += sum(0.04 for w in all_words if w in lib)
        name_hits = sum(1 for w in name_words if w in lib)
        s += name_hits * 0.50
        if name_words and name_hits == 0: s -= 0.60
        if vol_fourn:
            vol_cm = extract_volume(p['libelle'])
            if vol_cm:
                if vol_cm == vol_fourn: s += 0.30
                elif abs(vol_cm - vol_fourn) <= 5: s += 0.05
                else: s -= 0.30
            else: s -= 0.10
        if s > best_score:
            best_score, best = s, p
    return best, round(best_score, 2)

# ── Parsers ────────────────────────────────────────────────────────────────────

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
                if re.match(r'#REF\d+-\d+', nl) or re.match(r'^Page\s*:', nl) or 'Colisage' in nl: break
                if re.match(r'^\d{1,3}$', nl) and j > i + 1: break
                desc_parts.append(nl); j += 1
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

def parse_lca(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    i = 0
    while i < len(lines):
        if re.match(r'^#REF\d+-\d+$', lines[i]):
            ref = lines[i]
            desc_parts = []
            j = i + 1
            while j < len(lines):
                nl = lines[j]
                if re.match(r'^#REF\d+-\d+$', nl): break
                if re.match(r'^\d+$', nl) and j > i + 1: break
                if any(x in nl for x in ['Sous-total','RESERVE','Aucun','IBAN','Base HT']): break
                desc_parts.append(nl); j += 1
            desc = ' '.join(desc_parts)
            qty = 0
            if j < len(lines) and re.match(r'^\d+$', lines[j]): qty = int(lines[j])
            nic_m = re.search(r'[Nn]icotine\s*:\s*(\d+)\s*mg', desc)
            nic = str(int(nic_m.group(1))) if nic_m else '0'
            produit = re.sub(r'\s*-\s*[Dd]osage\s+[Nn]icotine.*', '', desc)
            produit = re.sub(r'\s*-\s*[Cc]ontenance.*', '', produit)
            produit = re.sub(r'\s*-\s*[Cc]ouleur.*', '', produit)
            produit = re.sub(r'\s*\([^)]*\)', '', produit).strip()
            if qty > 0 and 'PLV' not in produit:
                key = ref + '|' + nic
                if key not in seen:
                    seen.add(key)
                    items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
            i = j + 1
        else:
            i += 1
    return items

def parse_lvp(text):
    has_tva_col = 'Code\nTVA' in text or 'Code TVA' in text
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    SKIP = {'Référence','Désignation','Quantité','PU HT','Montant HT','Code','TVA',
            'Sous-total HT','INCOTERM DAP','TOTAL','Base HT','Taux TVA','Montant TVA',
            'Total','Remise','IBAN','BIC','RESERVE','LVP DISTRIBUTION'}
    def is_ref_lvp(s):
        if len(s) < 4: return False
        if not re.match(r'^[A-Z0-9][A-Z0-9\-\.]{3,}$', s): return False
        if re.match(r'^\d+[\.,]\d+$', s): return False
        return s not in SKIP
    i = 0
    while i < len(lines):
        line = lines[i]
        if not is_ref_lvp(line): i += 1; continue
        ref = line; j = i + 1; desc_parts = []
        while j < len(lines):
            nl = lines[j]
            if is_ref_lvp(nl): break
            if re.match(r'^\d{1,2}$', nl): break
            if re.match(r'^\d+[\.,]\d+$', nl): break
            if any(w in nl.lower() for w in ['sous-total','reserve','incoterm','iban']): break
            if nl: desc_parts.append(nl)
            j += 1
        desc = ' '.join(desc_parts)
        nic_m = re.search(r'[Nn]icotine\s*[:\(]\s*(\d+)\s*mg', desc)
        if nic_m: nic = str(int(nic_m.group(1)))
        else:
            nic = '0'
            for k in range(i+1, min(i+8, len(lines))):
                nm = re.match(r'^(\d+)mg$', lines[k], re.I)
                if nm: nic = str(int(nm.group(1))); break
        qty = 0
        while j < len(lines):
            nl = lines[j]
            if re.match(r'^\d+$', nl):
                candidate = int(nl)
                if has_tva_col and candidate <= 2 and j+1 < len(lines) and re.match(r'^\d+$', lines[j+1]):
                    j += 1; qty = int(lines[j])
                else: qty = candidate
                break
            if is_ref_lvp(nl): break
            if re.match(r'^\d+[\.,]\d+$', nl): break
            j += 1
        produit = re.sub(r'\s*\([^)]*\)', '', desc)
        produit = re.sub(r'\s*-\s*[Nn]icotine.*', '', produit)
        produit = re.sub(r'\s*-\s*[Cc]ouleur.*', '', produit)
        produit = re.sub(r'\s*-\s*[Cc]ontenance.*', '', produit)
        produit = re.sub(r'\s*-\s*[Oo]hm.*', '', produit)
        produit = re.sub(r'\s*-\s*[Vv]aleur.*', '', produit)
        produit = re.sub(r'\s*-\s*[Vv]ersion.*', '', produit)
        produit = re.sub(r'\s+(0mg|\d+mg)$', '', produit)
        produit = re.sub(r'\s+', ' ', produit).strip()
        skip_words = ['offert','base 1l','booster 10ml','fiole','bobine','film','chargeur','rouleau']
        if qty > 0 and len(produit) > 5 and not any(w in produit.lower() for w in skip_words):
            key = ref + '|' + nic
            if key not in seen:
                seen.add(key)
                items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
        i = j + 1
    return items

def parse_greenvillage(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    HEADER_STOP = ['Adresse','ROMAIN','Vapo','Centre','France','Customer','Référence','Produit']
    i = 0
    while i < len(lines):
        line = lines[i]
        m = re.match(r'^(\d[A-Z0-9]{2,}[\-]?)$', line)
        if not m: i += 1; continue
        ref_part1 = line; j = i + 1; ref_part2 = ''
        if j < len(lines) and re.match(r'^[A-Z0-9]{3,}$', lines[j]) and lines[j] != '20 %':
            ref_part2 = lines[j]; j += 1
        ref = ref_part1 + ref_part2
        if any(h in ' '.join(lines[max(0,i-3):i]) for h in HEADER_STOP): i += 1; continue
        desc_parts = []
        while j < len(lines):
            nl = lines[j]
            if nl == '20 %': break
            if re.match(r'^\d[A-Z0-9]{2,}[\-]?$', nl): break
            if any(x in nl for x in ['Total produits','Détail','Greenvillage','Powered']): break
            desc_parts.append(nl); j += 1
        desc = ' '.join(desc_parts)
        if j < len(lines) and lines[j] == '20 %': j += 3
        qty = 0
        if j < len(lines) and re.match(r'^\d+$', lines[j]):
            qty = int(lines[j]); j += 1
        nic_m = re.search(r'[Dd]éclinaison\s*:\s*(\d+)\s*mg', desc)
        if not nic_m: nic_m = re.search(r'(\d+)\s*mg', desc)
        if nic_m: nic = '0' if int(nic_m.group(1)) == 0 else str(int(nic_m.group(1)))
        else: nic = '0'
        produit = re.sub(r'\s*-\s*[Dd]éclinaison.*', '', desc)
        produit = re.sub(r'\s*\([^)]*\)', '', produit)
        produit = re.sub(r'\s*\(Par \d+\)', '', produit, flags=re.I)
        produit = produit.strip()
        skip = ['adresse','romain','vapo','centre','france','customer','tank en pyrex','offert']
        if qty > 0 and len(produit) > 5 and not any(s in produit.lower() for s in skip):
            key = ref + '|' + nic
            if key not in seen:
                seen.add(key)
                items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
        i = j
    return items

def parse_greenvillage2(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.match(r'^[45][A-Z]{2,4}-[A-Z0-9]+$', line):
            ref = line; j = i + 1; desc = ''
            if j < len(lines) and not re.match(r'^\d', lines[j]) and not re.match(r'^[45][A-Z]', lines[j]):
                desc = lines[j]; j += 1
            qty = 0
            if j < len(lines) and re.match(r'^\d+,\d+$', lines[j]):
                try: qty = int(float(lines[j].replace(',','.')))
                except: pass
                j += 1
            while j < len(lines) and (re.match(r'^\d+[,\.]', lines[j]) or '%' in lines[j]): j += 1
            nic = '0'
            if j < len(lines):
                nic_m = re.search(r'(\d+)mg', lines[j])
                if nic_m:
                    nic = str(int(nic_m.group(1)))
                    if re.search(r'\d+[.,]\d+\s*(?:ohm|\?)', lines[j], re.I): desc += ' ' + lines[j]
                    j += 1
            produit = re.sub(r'^\(\w+\)\s*', '', desc)
            produit = re.sub(r'\s*\d+mg.*$', '', produit)
            produit = re.sub(r'\s*Concentré/.*$', '', produit)
            produit = re.sub(r'\s*DTL/.*$', '', produit).strip()
            skip = ['port', 'frais de port']
            if qty > 0 and len(produit) > 3 and not any(s in produit.lower() for s in skip):
                key = ref + '|' + nic
                if key not in seen:
                    seen.add(key)
                    items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
            i = j
        else:
            i += 1
    return items

def parse_gfc(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    i = 0
    while i < len(lines):
        line = lines[i]
        if not re.match(r'^GFC\d+', line): i += 1; continue
        ref = line; j = i + 1; desc_parts = []
        while j < len(lines):
            nl = lines[j]
            if re.match(r'^GFC\d+', nl): break
            if re.match(r'^\d{8,}$', nl): break
            if any(x in nl for x in ['Sous-total','RESERVE','Base HT','Total HT','GFC Provap']): break
            desc_parts.append(nl); j += 1
        desc = ' '.join(desc_parts)
        if j < len(lines) and re.match(r'^\d{8,}$', lines[j]): j += 1
        qty = 0
        if j < len(lines) and re.match(r'^\d+$', lines[j]): qty = int(lines[j]); j += 1
        nic_m = re.search(r'[Nn]icotine\s*:\s*(\d+)\s*mg', desc)
        nic = str(int(nic_m.group(1))) if nic_m else '0'
        couleur_m = re.search(r'[Cc]ouleur\s*:\s*(.+?)$', desc)
        couleur = couleur_m.group(1).strip() if couleur_m else ''
        ohm_m = re.search(r'valeur\s*:\s*([^,]+ohm|\d+[.,]\d+\s*(?:ohm|\?))', desc, re.I)
        ohm_str = ohm_m.group(1).strip() if ohm_m else ''
        produit = re.sub(r'\s*-\s*[Mm]od[eè]le.*', '', desc)
        produit = re.sub(r'\s*-\s*[Vv]aleur.*', '', produit)
        produit = re.sub(r'\s*-\s*[Cc]ouleur.*', '', produit)
        produit = re.sub(r'\s*-\s*[Tt]aille.*', '', produit)
        produit = re.sub(r'\s*\([^)]*\)', '', produit).strip()
        if couleur and any(x in produit.lower() for x in ['kit','pack','box','pod']):
            produit = produit + ' - Couleur : ' + couleur
        if ohm_str and any(x in produit.lower() for x in ['resistance','cartouche','resistances']):
            produit = produit + ' ' + ohm_str
        skip = ['fid','ornement','accu vtc','accu 50s','accu 18650','accu 21700']
        if qty > 0 and len(produit) > 3 and not any(s in produit.lower() for s in skip):
            key = ref + '|' + nic
            if key not in seen:
                seen.add(key)
                items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
        i = j
    return items

def parse_adns(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    i = 0
    while i < len(lines):
        line = lines[i]
        if i+1 < len(lines) and lines[i+1] in ('UE', 'HORS UE') and len(line) > 5:
            desc = line; j = i + 2
            qty = 0
            if j < len(lines) and re.match(r'^\d+$', lines[j]): qty = int(lines[j]); j += 1
            nic_m = re.search(r'[Dd]osage\s+nicotine\s*:\s*0*(\d+)\s*mg', desc)
            nic = str(int(nic_m.group(1))) if nic_m else '0'
            produit = re.sub(r'\s*-\s*[Dd]osage\s+nicotine.*', '', desc)
            produit = re.sub(r'\s*-\s*[Cc]ouleur.*', '', produit)
            produit = re.sub(r'\s*-\s*[Ii]ntensit.*', '', produit)
            produit = re.sub(r'\s*\([^)]*\)', '', produit).strip()
            skip = ['echantillon','flyer','reduction','goodies']
            if qty > 0 and len(produit) > 3 and not any(s in produit.lower() for s in skip):
                key = desc[:30] + '|' + nic
                if key not in seen:
                    seen.add(key)
                    items.append({'ref': 'ADNS', 'produit': produit, 'nic': nic, 'qty': qty})
            i = j
        else:
            i += 1
    return items

def parse_grossiste(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    i = 0
    while i < len(lines):
        line = lines[i]
        if not re.match(r'^AR\d{5,}$', line): i += 1; continue
        ref = line; j = i + 1; desc_parts = []
        while j < len(lines):
            nl = lines[j]
            if nl == '20 %': break
            if re.match(r'^AR\d{5,}$', nl): break
            if any(x in nl for x in ['Reductions','Détail','Grossiste','Powered']): break
            desc_parts.append(nl); j += 1
        desc = ' '.join(desc_parts)
        if j < len(lines) and lines[j] == '20 %':
            j += 1
            if j < len(lines) and (lines[j] == '--' or '\u20ac' in lines[j]): j += 1
            if j < len(lines) and '\u20ac' in lines[j]: j += 1
        qty = 0
        if j < len(lines) and re.match(r'^\d+$', lines[j]): qty = int(lines[j]); j += 1
        nic_m = re.search(r'[Dd]osage\s*:\s*(\d+)\s*mg', desc)
        nic = str(int(nic_m.group(1))) if nic_m else '0'
        produit = re.sub(r'\s*-\s*[Dd]osage\s*:.*', '', desc)
        produit = re.sub(r'\s*-\s*[Ss]aveur\s*:.*', '', produit)
        produit = re.sub(r'\s*-\s*[Cc]ouleur\s*:.*', '', produit)
        produit = re.sub(r'\s*\([Bb]oite\s+de\s+\d+\)', '', produit)
        produit = re.sub(r'\s*\([^)]*\)', '', produit)
        produit = re.sub(r'^E liquide\s+', '', produit, flags=re.I).strip()
        saveur_m = re.search(r'[Ss]aveur\s*:\s*(.+?)(?:\s*$|\s*-)', desc)
        if saveur_m: produit = produit + ' ' + saveur_m.group(1).strip()
        skip = ['reductions','offerte','livraison','1 boite']
        if qty > 0 and len(produit) > 3 and not any(s in produit.lower() for s in skip):
            key = ref + '|' + nic
            if key not in seen:
                seen.add(key)
                items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
        i = j
    return items

def parse_airmust(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    STOP = ['Detail des taxes','Total produits','Frais','Total (HT)','Total','Airmust',
            'Powered','Conditions','ATTENTION','Transporteur','Moyen']
    i = 0
    while i < len(lines):
        line = lines[i]
        if re.match(r'^\d{10}$', line) and i+1 < len(lines) and re.match(r'^\d{3}$', lines[i+1]):
            ref = line + lines[i+1]; j = i + 2; desc_parts = []
            while j < len(lines):
                nl = lines[j]
                if nl == '20 %': break
                if re.match(r'^\d{10}$', nl): break
                if any(x in nl for x in STOP): break
                desc_parts.append(nl); j += 1
            desc = ' '.join(desc_parts)
            if j < len(lines) and lines[j] == '20 %':
                j += 1
                if j < len(lines) and '\u20ac' in lines[j]: j += 1
                if j < len(lines) and '\u20ac' in lines[j]: j += 1
            qty = 0
            if j < len(lines) and re.match(r'^\d+$', lines[j]): qty = int(lines[j]); j += 1
            nic_m = re.search(r'[Nn]icotine\s*:\s*(\d+)\s*mg', desc)
            nic = str(int(nic_m.group(1))) if nic_m else '0'
            produit = re.sub(r'\s*\([Nn]icotine.*?\)', '', desc)
            produit = re.sub(r'^(AIRMUST|Paperland|Ferox|Le Primeur Fresh)\s*[\u2022\xb7]\s*', '', produit, flags=re.I)
            produit = produit.strip()
            if qty > 0 and len(produit) > 2:
                key = ref + '|' + nic
                if key not in seen:
                    seen.add(key)
                    items.append({'ref': ref, 'produit': produit, 'nic': nic, 'qty': qty})
            i = j
        else:
            i += 1
    return items

def parse_ag(text):
    items, seen = [], set()
    lines = [l.strip() for l in text.split('\n')]
    STOP = ['TOTAL','Taux','TVA','CREDIT','IBAN','BIC','AG Consulting','Shipping',
            'Mode de','Date','Ref.','Montant','Description','NF525','Page','FRA',
            'Email','Tel','Mathieu','Vapochill','La Grande','Saunay','Romain',
            'PAIMENT','VAPOCHILL']
    CBD_KEYWORDS = ['B20','B52','THV2','CBD','FILTRE','Hash','Biscotti','Gorilla',
                    'Papaya','Banana','Amnesia','Widow','Blueberry','Cannatonic',
                    'Jack','Apple','Watermelon','Gary Payton','Mooncake','Recharge']
    i = 0
    while i < len(lines):
        line = lines[i]
        is_cbd = any(x in line for x in CBD_KEYWORDS)
        if len(line) > 5 and is_cbd and not any(x in line for x in STOP):
            desc_parts = [line]; j = i + 1
            while j < len(lines):
                nl = lines[j]
                if re.match(r'^\d+$', nl): break
                if re.match(r'^\d+[\.,]\d+$', nl): break
                if any(x in nl for x in STOP): break
                if not nl: break
                desc_parts.append(nl); j += 1
            desc = ' '.join(desc_parts)
            qty = 0
            if j < len(lines) and re.match(r'^\d+$', lines[j]): qty = int(lines[j])
            if qty > 0 and 'Shipping' not in desc:
                produit = re.sub(r'\s*-\s*Sommités.*', '', desc)
                produit = re.sub(r'\s*-\s*Lot\s*n[°o].*', '', produit)
                produit = re.sub(r'\s*\(THC.*?\)', '', produit)
                produit = re.sub(r'\s*-\s*GTIN.*', '', produit)
                produit = re.sub(r'\s*\(THV.*?\)', '', produit)
                produit = produit.strip()
                key = produit[:40] + '|0'
                if key not in seen:
                    seen.add(key)
                    items.append({'ref': 'AG-CBD', 'produit': produit, 'nic': '0', 'qty': qty})
            i = j + 1 if qty > 0 else i + 1
        else:
            i += 1
    return items

def parse_facture_standard(text):
    items = []
    for m in re.finditer(
        r'([A-Z0-9]{6,})\s+(.+?)(?:Nicotine\s*:\s*(\d+)\s*mg[^)]*\)?\s*)?(?:0\s*%|20\s*%)\s+[\d,]+\s*\u20ac\s+(\d+)',
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
    if 'AG Consulting' in text:
        return parse_ag(text)
    if 'Airmust' in text or 'AIRMUST' in text:
        return parse_airmust(text)
    if 'Grossiste Ecigarette' in text or 'grossiste-ecigarette' in text.lower():
        return parse_grossiste(text)
    if 'GFC Provap' in text or 'gfc-provap' in text.lower():
        return parse_gfc(text)
    if 'ADNS' in text and 'Vente en gros' in text:
        return parse_adns(text)
    if 'LCA DISTRIBUTION' in text:
        return parse_lca(text)
    if 'LVP DISTRIBUTION' in text:
        return parse_lvp(text)
    if 'greenvillage' in text.lower() or 'green village' in text.lower():
        if 'Alliance Distribution' in text or 'Px net u.' in text:
            return parse_greenvillage2(text)
        return parse_greenvillage(text)
    if '#REF' in text and ('Colisage' in text or 'Dosage Nicotine' in text):
        return parse_bl(text)
    return parse_facture_standard(text)

# ── Génération Excel ───────────────────────────────────────────────────────────
def generate_excel(results, filename):
    VERT="C6EFCE"; VERT_TXT="276221"; ORANGE="FFEB9C"; ORANGE_TXT="9C5700"
    BLEU="1F4E79"; BLEU2="2E75B6"
    thin = Side(style='thin', color='CCCCCC')
    border = Border(left=thin,right=thin,top=thin,bottom=thin)
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "Entree stock"
    ws.merge_cells('A1:H1')
    ws['A1'] = f"ENTREE DE STOCK - {filename}"
    ws['A1'].font = Font(bold=True,size=12,color='FFFFFF',name='Arial')
    ws['A1'].fill = PatternFill('solid',start_color=BLEU)
    ws['A1'].alignment = Alignment(horizontal='center',vertical='center')
    ws.row_dimensions[1].height = 22
    ok_c = sum(1 for r in results if r['statut']=='OK')
    ws.merge_cells('A2:H2')
    ws['A2'] = f"{len(results)} references | OK: {ok_c} | A verifier: {len(results)-ok_c}"
    ws['A2'].font = Font(size=10,color='595959',name='Arial')
    ws['A2'].fill = PatternFill('solid',start_color='DEEAF1')
    ws['A2'].alignment = Alignment(horizontal='center',vertical='center')
    ws.row_dimensions[2].height = 15
    for col, h in enumerate(['#','Ref Fournisseur','Produit Fournisseur','Nic/Ohm','Qte',
                              'Libelle Cash Mag','ID Cash Mag','Statut'], 1):
        c = ws.cell(row=3,column=col,value=h)
        c.font = Font(bold=True,color='FFFFFF',name='Arial',size=10)
        c.fill = PatternFill('solid',start_color=BLEU2)
        c.alignment = Alignment(horizontal='center',vertical='center')
        c.border = border
    ws.row_dimensions[3].height = 18
    for i, r in enumerate(results, 1):
        rn = i+3; is_av = r['statut']=='A VERIFIER'
        for col, val in enumerate([i,r['ref'],r['produit'],r['nic'],r['qty'],
                                    r['cashMagLibelle'],r['cashMagId'],r['statut']], 1):
            c = ws.cell(row=rn,column=col,value=val)
            c.font = Font(name='Arial',size=10); c.border = border
            c.alignment = Alignment(vertical='center',horizontal='center' if col in [1,4,5,7,8] else 'left')
            if col == 8:
                c.fill = PatternFill('solid',start_color=ORANGE if is_av else VERT)
                c.font = Font(name='Arial',size=10,bold=True,color=ORANGE_TXT if is_av else VERT_TXT)
        ws.row_dimensions[rn].height = 16
    for col, w in zip('ABCDEFGH',[4,18,36,10,6,40,13,13]):
        ws.column_dimensions[col].width = w
    output = io.BytesIO()
    wb.save(output); output.seek(0)
    return output

# ── Routes ─────────────────────────────────────────────────────────────────────
RESULTATS_CACHE = {}

@app.route('/')
def index():
    return render_template('index.html', nb_produits=len(CATALOGUE))

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier recu'}), 400
    file = request.files['file']
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Envoyez un fichier PDF'}), 400
    try:
        doc = fitz.open(stream=file.read(), filetype='pdf')
        text = "\n".join(page.get_text() for page in doc)
        items = parse_doc(text)
        if not items:
            return jsonify({'error': 'Aucune reference trouvee. Fournisseur non supporte ?'}), 400
        results = []
        for item in items:
            ohm = extract_ohm(item['produit'])
            is_resistance = any(kw in item['produit'].lower() for kw in
                ['résistance','resistance','cartouche','clearomiseur'])
            match, score = find_best(item['produit'], item['nic'])
            nic_display = item['nic']
            if is_resistance and ohm:
                nic_display = ohm + '\u03a9'
            elif nic_display not in ('0','00'):
                nic_display = nic_display + 'mg'
            else:
                nic_display = '0mg'
            results.append({**item,
                'nic': nic_display,
                'cashMagLibelle': match['libelle'] if match else 'NON TROUVE',
                'cashMagId': str(match['id']) if match else '',
                'score': score,
                'statut': 'OK' if score >= 0.65 else 'A VERIFIER'})
        fname = os.path.splitext(file.filename)[0]
        cache_id = str(uuid.uuid4())
        RESULTATS_CACHE[cache_id] = {'results': results, 'fname': fname}
        ok = sum(1 for r in results if r['statut']=='OK')
        return jsonify({'cache_id': cache_id, 'fname': fname, 'total': len(results),
                        'ok': ok, 'av': len(results)-ok, 'results': results})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<cache_id>')
def download(cache_id):
    if cache_id not in RESULTATS_CACHE:
        return jsonify({'error': 'Session expiree, retraitez le PDF'}), 404
    data = RESULTATS_CACHE[cache_id]
    excel = generate_excel(data['results'], data['fname'])
    return send_file(excel, as_attachment=True,
        download_name=f'entree_stock_{data["fname"]}.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
