"""
DDC Generator V4F — Backend FastAPI
Mise en forme 1:1 avec le template V4F
"""
import io, json, os, re, zipfile, copy, uuid, unicodedata
from datetime import datetime
from functools import lru_cache
from pathlib import Path
import pdfplumber, ollama
from lxml import etree
from fastapi import FastAPI, File, UploadFile, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
try:
    from pro_builder import build_pro_pptx
except ImportError:
    from backend.pro_builder import build_pro_pptx

# ── CONFIG ────────────────────────────────────────────────────────────────────
PASSWORD      = os.getenv("DDC_PASSWORD", "v4f2025")
MODEL         = os.getenv("OLLAMA_MODEL", "mistral-nemo")
OLLAMA_HOST   = os.getenv("OLLAMA_HOST",  "http://localhost:11434")
TEMPLATE_PATH = Path(__file__).parent.parent / "TEMPLATE.pptx"
OUTPUT_DIR    = Path(__file__).parent.parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)
HISTORY_FILE  = OUTPUT_DIR / "history.json"

def _load_history():
    if HISTORY_FILE.exists():
        try:
            return json.loads(HISTORY_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return []

def _save_history(history):
    HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding='utf-8')

history = _load_history()

app = FastAPI(title="DDC Generator V4F")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

# ── AUTH ──────────────────────────────────────────────────────────────────────
class LoginRequest(BaseModel):
    password: str

@app.post("/api/login")
def login(req: LoginRequest):
    if req.password == PASSWORD:
        return {"ok": True, "token": "v4f-token-" + PASSWORD}
    raise HTTPException(401, "Mot de passe incorrect")

def check_auth(request: Request):
    token = request.headers.get("Authorization","").replace("Bearer ","")
    if token != "v4f-token-" + PASSWORD:
        raise HTTPException(401, "Non autorisé")

# ── NAMESPACES ────────────────────────────────────────────────────────────────
A  = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P  = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
RN = 'http://schemas.openxmlformats.org/package/2006/relationships'
EN = '\u2013'  # en-dash
BL = '002A6D'
CY = '43CEFF'

# ── LAYOUT — calibré sur mesures exactes du template V4F ─────────────────────
# Depuis analyse pixel-perfect du template original:
# Template V4F de référence: GAE ateba christian - DDC V4F.pptx
GAP_GRP   = 40000
TEMPLATE_MISSION_GAP = 120000
SLIDE_CONTENT_LIMIT = 9694442
MAX_MISSIONS_PER_SLIDE = 3
MAX_MISSIONS_TOTAL = 24
MAX_REALISATIONS_PER_MISSION = 5
REALISATION_MAX_WORDS = 18
REALISATION_MAX_CHARS = 140
THREE_MISSION_MAX_FILL = 0.88
TWO_MISSION_TARGET_FILL = 0.84
LAYOUT_PROFILES = [
    {
        'name': 'normal',
        'bullet_h': 255000,
        'bullet_overhead': 65000,
        'group_h': 323165,
        'title_line_h': 175000,
        'title_char_limit': 60,
        'bullet_char_limit': 86,
        'bullet_font_sz': 1600,
        'bullet_line_spc': 108000,
        'bullet_spc_aft': 0,
        'title_font_sz': 1500,
        'sep_font_sz': 1500,
        'title_kern': 1000,
    },
    {
        'name': 'compact',
        'bullet_h': 235000,
        'bullet_overhead': 50000,
        'group_h': 305000,
        'title_line_h': 160000,
        'title_char_limit': 66,
        'bullet_char_limit': 96,
        'bullet_font_sz': 1500,
        'bullet_line_spc': 104000,
        'bullet_spc_aft': 0,
        'title_font_sz': 1425,
        'sep_font_sz': 1425,
        'title_kern': 1000,
    },
    {
        'name': 'dense',
        'bullet_h': 215000,
        'bullet_overhead': 40000,
        'group_h': 292000,
        'title_line_h': 145000,
        'title_char_limit': 72,
        'bullet_char_limit': 108,
        'bullet_font_sz': 1400,
        'bullet_line_spc': 100000,
        'bullet_spc_aft': 0,
        'title_font_sz': 1350,
        'sep_font_sz': 1350,
        'title_kern': 900,
    },
    {
        'name': 'airy',
        'bullet_h': 205000,
        'bullet_overhead': 35000,
        'group_h': 276000,
        'title_line_h': 135000,
        'title_char_limit': 78,
        'bullet_char_limit': 118,
        'bullet_font_sz': 1350,
        'bullet_line_spc': 98000,
        'bullet_spc_aft': 0,
        'title_font_sz': 1275,
        'sep_font_sz': 1275,
        'title_kern': 800,
    },
]

# Positions fixes du template (GAE V4F de référence)
GRP1_Y   = 2516403
ZONE1_Y  = 3151599
ZONE1_TOP_OFFSET = ZONE1_Y - GRP1_Y  # offset entre groupe et zone bullets
TEMPLATE_MEMORY = {
    'Titre 1': {'x': 241199, 'y': 1366942, 'cx': 5684115, 'cy': 1437784},
    'Espace réservé du texte 13': {'x': 6741548, 'y': 476108, 'cx': 9298524, 'cy': 369332},
    'Espace réservé du texte 5': {'x': 7124921, 'y': 1046527, 'cx': 10921880, 'cy': 1272847},
    'Espace réservé du texte 2': {'x': 120599, 'y': 3020815, 'cx': 5784217, 'cy': 2206191},
    'Espace réservé du texte 3': {'x': 241199, 'y': 5230037, 'cx': 4859999, 'cy': 2349182},
    'Espace réservé du texte 4': {'x': 241199, 'y': 7592902, 'cx': 4860000, 'cy': 2206191},
    'Groupe 6': {'x': 6741548, 'y': 2516403, 'cx': 10617312, 'cy': 323165},
    'ZoneTexte 14': {'x': 7344597, 'y': 1443789, 'cx': 10315605, 'cy': 323165},
    'ZoneTexte 15': {'x': 6287026, 'y': 3151599, 'cx': 11759775, 'cy': 1477328},
    'Groupe 9': {'x': 6741548, 'y': 5415425, 'cx': 8762873, 'cy': 323165},
    'ZoneTexte 16': {'x': 7380694, 'y': 6433983, 'cx': 8423414, 'cy': 323165},
    'ZoneTexte 18': {'x': 6321169, 'y': 5909094, 'cx': 11759775, 'cy': 1477328},
    'Groupe 11': {'x': 6741548, 'y': 8313613, 'cx': 8762873, 'cy': 323165},
    'ZoneTexte 17': {'x': 7380694, 'y': 6433983, 'cx': 8423414, 'cy': 323165},
    'ZoneTexte 19': {'x': 6321168, 'y': 8786501, 'cx': 11759775, 'cy': 907941},
}
TEXT_SHAPES = set(TEMPLATE_MEMORY)
MISSION_TITLE_NAMES = ['ZoneTexte 14', 'ZoneTexte 16', 'ZoneTexte 17']
MISSION_ZONE_NAMES = ['ZoneTexte 15', 'ZoneTexte 18', 'ZoneTexte 19']
GROUP_NAMES = ['Groupe 6', 'Groupe 9', 'Groupe 11']
# Standalone shapes that duplicate group children — must be hidden to avoid placeholder bleed
STANDALONE_DUPLICATES = {'ZoneTexte 14', 'ZoneTexte 16', 'ZoneTexte 17', 'ZoneTexte 22'}
EDITABLE_TEXT_SHAPES = {
    'Titre 1',
    'Espace réservé du texte 5',
    'Espace réservé du texte 2',
    'Espace réservé du texte 3',
    'Espace réservé du texte 4',
    *MISSION_TITLE_NAMES,
    *MISSION_ZONE_NAMES,
}
STYLE_MEMORY = {
    'mission_title_font': 'Montserrat',
    'mission_title_color': BL,
    'mission_bullet_font': 'Calibri',
    'mission_bullet_color': 'black',
}

MONTH_NAME_TO_NUM = {
    'janvier': 1, 'jan': 1, 'january': 1,
    'fevrier': 2, 'février': 2, 'fev': 2, 'fév': 2, 'february': 2, 'feb': 2,
    'mars': 3, 'march': 3,
    'avril': 4, 'avr': 4, 'april': 4, 'apr': 4,
    'mai': 5, 'may': 5,
    'juin': 6, 'june': 6,
    'juillet': 7, 'juil': 7, 'july': 7, 'jul': 7,
    'aout': 8, 'août': 8, 'august': 8, 'aug': 8,
    'septembre': 9, 'sept': 9, 'sep': 9, 'september': 9,
    'octobre': 10, 'oct': 10, 'october': 10,
    'novembre': 11, 'nov': 11, 'november': 11,
    'decembre': 12, 'décembre': 12, 'dec': 12, 'déc': 12, 'december': 12,
}
MONTH_PATTERN = '|'.join(sorted((re.escape(k) for k in MONTH_NAME_TO_NUM), key=len, reverse=True))
DATE_TOKEN_RE = re.compile(
    rf"(?:(?:\d{{1,2}}[\/.\-]\d{{2,4}})|(?:\d{{1,2}}\s+)?(?:{MONTH_PATTERN})\s+\d{{2,4}}|(?:present|présent|current|now|aujourd'hui))",
    re.IGNORECASE,
)

def normalize_text(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip()

ABBREVIATIONS = [
    ('Contrôle', 'Ctrl'),
    ('conformité', 'conf.'),
    ('Management', 'Mgmt'),
    ('Professionnelle', 'Prof.'),
    ('Supérieure', 'Sup.'),
    ('Formation', 'Form.'),
    ('Gestion', 'Gest.'),
    ('spécialité', 'spé.'),
    ('Spécialiste', 'Spéc.'),
    ('Technique', 'Tech.'),
    ('Assurances', 'Assur.'),
    ('Assurance', 'Assur.'),
    ('Analyste', 'Analyste'),
]

def apply_abbreviations(text):
    out = normalize_text(text)
    for src, dst in ABBREVIATIONS:
        out = re.sub(rf'\b{re.escape(src)}\b', dst, out, flags=re.IGNORECASE)
    out = out.replace(' et ', ' & ')
    return normalize_text(out)

def trim_words(text, max_chars):
    text = normalize_text(text)
    if len(text) <= max_chars:
        return text
    words = text.split()
    out = []
    for word in words:
        candidate = ' '.join(out + [word])
        if len(candidate) > max_chars:
            break
        out.append(word)
    if not out:
        return text[:max_chars].rstrip()
    return ' '.join(out)

def smart_fit(text, max_chars):
    text = normalize_text(text)
    if len(text) <= max_chars:
        return text
    abbreviated = apply_abbreviations(text)
    if len(abbreviated) <= max_chars:
        return abbreviated
    return trim_words(abbreviated, max_chars)

def fit_formation(text):
    text = smart_fit(text, 100)
    parts = [normalize_text(part) for part in re.split(r'\s+-\s+', text) if normalize_text(part)]
    if len(text) <= 100 or len(parts) < 2:
        return text
    left = smart_fit(parts[0], 44)
    right = smart_fit(' - '.join(parts[1:]), 52)
    merged = f'{left} - {right}'
    return trim_words(merged, 100)

def fit_title_field(text, max_chars):
    return smart_fit(text, max_chars)

TRAILING_FILLER_WORDS = {
    '&', 'et', 'ou', 'de', 'des', 'du', 'la', 'le', 'les', 'pour', 'par',
    'aux', 'au', 'avec', 'dans', 'sur', 'en', 'd', 'l',
}

def trim_trailing_fillers(text):
    words = normalize_text(text).split()
    while words:
        last = re.sub(r"[^\w&]+$", "", ascii_fold(words[-1]))
        if last in TRAILING_FILLER_WORDS or not last:
            words.pop()
            continue
        break
    return ' '.join(words)

def fit_realisation(text, max_words=REALISATION_MAX_WORDS, max_chars=REALISATION_MAX_CHARS):
    text = normalize_text(text)
    if not text:
        return ''
    text = re.sub(r'\s*\(([^)]{1,18})\)\s*', r' (\1) ', text)
    original = normalize_text(text)
    words = original.split()
    # Préserver le wording du CV tant que ça reste raisonnablement affichable.
    if len(words) <= 24 and len(original) <= 180:
        return original

    text = apply_abbreviations(original)
    text = normalize_text(text)
    words = text.split()
    if len(words) <= max_words and len(text) <= max_chars:
        return text

    best = ' '.join(words[:max_words])
    best = trim_trailing_fillers(best)

    cut_points = [',', ';', ':']
    for sep in cut_points:
        parts = [normalize_text(part) for part in text.split(sep) if normalize_text(part)]
        if not parts:
            continue
        candidate = trim_trailing_fillers(parts[0])
        if candidate and len(candidate.split()) <= max_words and len(candidate) <= max_chars:
            best = candidate
            break

    if len(best.split()) > max_words or len(best) > max_chars:
        best = trim_words(best, max_chars)
        best = ' '.join(best.split()[:max_words])
        best = trim_trailing_fillers(best)

    if len(best.split()) < 5:
        fallback = ' '.join(words[:max_words])
        fallback = trim_trailing_fillers(fallback)
        if fallback:
            best = fallback

    return normalize_text(best)

def ascii_fold(text):
    text = normalize_text(text).lower()
    return ''.join(c for c in unicodedata.normalize('NFD', text) if unicodedata.category(c) != 'Mn')

def source_tokens(text):
    return {tok for tok in re.findall(r'[a-z0-9]{3,}', ascii_fold(text))}

def overlap_ratio(text, source_tok):
    toks = source_tokens(text)
    if not toks:
        return 0.0
    return len(toks & source_tok) / len(toks)

def grounded(text, source_text, source_tok=None, source_folded=None, min_ratio=0.6, min_hits=2):
    folded = ascii_fold(text)
    if not folded:
        return False
    if source_folded is None:
        source_folded = ascii_fold(source_text)
    if folded in source_folded:
        return True
    toks = source_tokens(text)
    if source_tok is None:
        source_tok = source_tokens(source_text)
    hits = len(toks & source_tok)
    return hits >= min_hits and (len(toks & source_tok) / len(toks)) >= min_ratio

def richer_text(candidate, reference):
    cand = normalize_text(candidate)
    ref = normalize_text(reference)
    if not cand:
        return False
    if not ref:
        return True
    return len(cand) > len(ref) and (cand.startswith(ref) or ref.startswith(cand) or cand[:24] == ref[:24])

def estimate_lines(text, max_chars):
    words = normalize_text(text).split()
    if not words:
        return 1
    lines, current = 1, 0
    for word in words:
        word_len = len(word)
        if current == 0:
            current = word_len
            continue
        if current + 1 + word_len <= max_chars:
            current += 1 + word_len
        else:
            lines += 1
            current = word_len
    return max(lines, 1)

def compact_date_range(text):
    return re.sub(r'\s*-\s*', f'{EN}', normalize_text(text))

def parse_month_year(token, is_end=False):
    token = normalize_text(token).lower()
    token = token.replace('present', 'présent')
    if token in {'présent', 'current', 'now', "aujourd'hui"}:
        now = datetime.now()
        return now.year, now.month
    m = re.search(r'(\d{1,2})[\/.\-](\d{4})', token)
    if m:
        month = int(m.group(1))
        year = int(m.group(2))
        if 1 <= month <= 12:
            return year, month
    m = re.search(r'(\d{1,2})[\/.\-](\d{2})(?!\d)', token)
    if m:
        month = int(m.group(1))
        year = int(m.group(2))
        if 1 <= month <= 12:
            current_two_digits = datetime.now().year % 100
            year += 2000 if year <= current_two_digits + 1 else 1900
            return year, month
    folded = ascii_fold(token)
    m = re.search(rf'(?:(\d{{1,2}})\s+)?({MONTH_PATTERN})\s+(\d{{2,4}})', folded, flags=re.IGNORECASE)
    if m:
        month_name = m.group(2).lower()
        month = MONTH_NAME_TO_NUM.get(month_name)
        year = int(m.group(3))
        if year < 100:
            current_two_digits = datetime.now().year % 100
            year += 2000 if year <= current_two_digits + 1 else 1900
        if month:
            return year, month
    m = re.search(r'(\d{4})', token)
    if m:
        year = int(m.group(1))
        return year, 12 if is_end else 1
    return None

def parse_date_range(text):
    text = normalize_text(text)
    if not text:
        return None
    tokens = [m.group(0) for m in DATE_TOKEN_RE.finditer(text)]
    if len(tokens) >= 2:
        start = parse_month_year(tokens[0], is_end=False)
        end = parse_month_year(tokens[-1], is_end=True)
    else:
        parts = re.split(r'\s*[–—-]\s*', text)
        if len(parts) < 2:
            return None
        start = parse_month_year(parts[0], is_end=False)
        end = parse_month_year(parts[-1], is_end=True)
    if not start or not end:
        return None
    start_idx = start[0] * 12 + (start[1] - 1)
    end_idx = end[0] * 12 + (end[1] - 1)
    now = datetime.now()
    now_idx = now.year * 12 + (now.month - 1)
    if start_idx > now_idx:
        return None
    end_idx = min(end_idx, now_idx)
    if end_idx < start_idx:
        return None
    return start_idx, end_idx

def derive_annees_xp(missions):
    ranges = []
    for mission in missions:
        parsed = parse_date_range(mission.get('duree'))
        if parsed:
            ranges.append(parsed)
    if not ranges:
        return ''
    ranges.sort()
    merged = [list(ranges[0])]
    for start, end in ranges[1:]:
        last = merged[-1]
        if start <= last[1] + 1:
            last[1] = max(last[1], end)
        else:
            merged.append([start, end])
    total_months = sum(end - start + 1 for start, end in merged)
    years = total_months // 12
    if years >= 2:
        return f'{years} ans'
    if years == 1:
        return '1 an'
    return f'{total_months} mois' if total_months else ''

def format_month_duration(total_months):
    if total_months <= 0:
        return ''
    if total_months < 12:
        return f'{total_months} mois'
    years = total_months // 12
    months = total_months % 12
    year_label = '1 an' if years == 1 else f'{years} ans'
    if months <= 0:
        return year_label
    return f'{year_label} {months} mois'

def display_mission_duration(text):
    raw = normalize_text(text)
    if not raw:
        return ''
    parsed = parse_date_range(raw)
    if parsed:
        start_idx, end_idx = parsed
        return format_month_duration(end_idx - start_idx + 1)
    months_match = re.search(r'(\d+)\s*mois', raw, flags=re.IGNORECASE)
    years_match = re.search(r'(\d+)\s*an[s]?', raw, flags=re.IGNORECASE)
    if years_match or months_match:
        years = int(years_match.group(1)) if years_match else 0
        months = int(months_match.group(1)) if months_match else 0
        return format_month_duration(years * 12 + months)
    return raw

def extract_date_tokens(text):
    return [m.group(0) for m in DATE_TOKEN_RE.finditer(normalize_text(text))]

def extract_experience_section(cv_text):
    text = str(cv_text or '')
    start_match = re.search(r'EXPERIENCE[S]?\s+PROFESSIONNELLE[S]?', text, flags=re.IGNORECASE)
    if not start_match:
        return ''
    section = text[start_match.end():]
    end_markers = [
        r'CERTIFICATIONS?(?:\s+ET\s+TRAINING)?',
        r'FORMATIONS?\s+COMPLEMENTAIRES?',
        r'CENTRES?\s+D[’\' ]INTERET',
        r'LANGUES?',
    ]
    end_positions = []
    for marker in end_markers:
        m = re.search(marker, section, flags=re.IGNORECASE)
        if m:
            end_positions.append(m.start())
    if end_positions:
        section = section[:min(end_positions)]
    return section

def looks_like_mission_header(line):
    text = normalize_text(line)
    if not text:
        return False
    tokens = extract_date_tokens(text)
    if len(tokens) < 2:
        return False
    folded = ascii_fold(text)
    if folded.startswith(('certification', 'training', 'langue')):
        return False
    if text.startswith('('):
        return False
    return True

def parse_mission_header(line):
    text = normalize_text(line)
    matches = list(DATE_TOKEN_RE.finditer(text))
    if len(matches) < 2:
        return None
    date_text = compact_date_range(f"{matches[0].group(0)} - {matches[-1].group(0)}")
    head = normalize_text(text[:matches[0].start()]).strip(' .,-–—/')
    head = re.sub(r'\bdepuis\b$', '', head, flags=re.IGNORECASE).strip(' .,-–—/')
    if not head:
        return None

    company = head
    poste = ''
    dash_parts = re.split(r'\s+[–—-]\s+', head, maxsplit=1)
    if len(dash_parts) == 2:
        company, poste = dash_parts[0], dash_parts[1]
    else:
        m = re.match(r'^(.*?\([^)]*\))\s+(.*)$', head)
        if m:
            company, poste = m.group(1), m.group(2)
    company = normalize_text(company).strip(' .,-–—/')
    poste = normalize_text(poste)
    poste = re.sub(r'\bdepuis\b\s*$', '', poste, flags=re.IGNORECASE)
    poste = poste.rstrip(' .,-–—/(').strip()
    return {
        'entreprise': company,
        'poste': poste,
        'duree': date_text,
        'realisations': [],
    }

def looks_like_date_only_line(line):
    text = normalize_text(line)
    if not text:
        return False
    folded = ascii_fold(text)
    stripped = DATE_TOKEN_RE.sub('', folded)
    stripped = re.sub(r'[\s().,/–—-]+', '', stripped)
    return bool(extract_date_tokens(text)) and not stripped

def preprocess_experience_lines(lines):
    prepared = []
    i = 0
    while i < len(lines):
        line = normalize_text(lines[i])
        if (
            i + 1 < len(lines)
            and not looks_like_mission_header(line)
            and looks_like_date_only_line(lines[i + 1])
        ):
            prepared.append(normalize_text(f'{line} {lines[i + 1]}'))
            i += 2
            continue
        prepared.append(line)
        i += 1
    return prepared

def merge_wrapped_lines(lines):
    merged = []
    for raw in lines:
        line = normalize_text(raw)
        if not line:
            continue
        if not merged:
            merged.append(line)
            continue
        prev = merged[-1]
        if (
            prev.endswith((',', '/', '-', EN))
            or line[:1].islower()
            or line.startswith((')', ']', '}', ';', ':'))
        ):
            merged[-1] = normalize_text(f'{prev} {line}')
        else:
            merged.append(line)
    return merged

def recover_missions_from_cv_text(cv_text):
    section = extract_experience_section(cv_text)
    if not section:
        return []
    lines = preprocess_experience_lines([normalize_text(line) for line in section.splitlines() if normalize_text(line)])
    recovered = []
    current = None
    body = []
    for line in lines:
        if looks_like_mission_header(line):
            if current:
                current['realisations'] = merge_wrapped_lines(body)[:MAX_REALISATIONS_PER_MISSION]
                recovered.append(current)
            current = parse_mission_header(line)
            body = []
            continue
        if current is not None:
            body.append(line)
    if current:
        current['realisations'] = merge_wrapped_lines(body)[:MAX_REALISATIONS_PER_MISSION]
        recovered.append(current)
    return recovered[:MAX_MISSIONS_TOTAL]

def same_mission(a, b):
    a_ent = ascii_fold(a.get('entreprise'))
    b_ent = ascii_fold(b.get('entreprise'))
    if not a_ent or not b_ent:
        return False
    company_match = a_ent in b_ent or b_ent in a_ent
    if not company_match:
        return False
    a_dur = ascii_fold(a.get('duree'))
    b_dur = ascii_fold(b.get('duree'))
    if a_dur and b_dur:
        return a_dur == b_dur or a_dur in b_dur or b_dur in a_dur
    a_poste = ascii_fold(a.get('poste'))
    b_poste = ascii_fold(b.get('poste'))
    if a_poste and b_poste:
        return a_poste in b_poste or b_poste in a_poste
    return False

def merge_recovered_missions(data, cv_text):
    recovered = recover_missions_from_cv_text(cv_text)
    if not recovered:
        return data
    base = copy.deepcopy(data) if isinstance(data, dict) else {}
    base_missions = base.get('missions', []) if isinstance(base.get('missions'), list) else []
    merged = []
    used = set()
    for rec in recovered:
        match_idx = None
        for idx, mission in enumerate(base_missions):
            if idx in used:
                continue
            if isinstance(mission, dict) and same_mission(mission, rec):
                match_idx = idx
                break
        if match_idx is None:
            merged.append(rec)
            continue
        used.add(match_idx)
        mission = copy.deepcopy(base_missions[match_idx])
        for field in ['entreprise', 'poste', 'duree']:
            if not normalize_text(mission.get(field)) and normalize_text(rec.get(field)):
                mission[field] = rec.get(field)
        rec_reals = [normalize_text(x) for x in rec.get('realisations', []) if normalize_text(x)]
        base_reals = [normalize_text(x) for x in mission.get('realisations', []) if normalize_text(x)]
        combined = []
        for item in rec_reals + base_reals:
            if item and not any(ascii_fold(item) == ascii_fold(existing) for existing in combined):
                combined.append(item)
        mission['realisations'] = combined[:MAX_REALISATIONS_PER_MISSION]
        merged.append(mission)
    for idx, mission in enumerate(base_missions):
        if idx not in used and isinstance(mission, dict):
            merged.append(copy.deepcopy(mission))
    base['missions'] = merged[:MAX_MISSIONS_TOTAL]
    return base

def mission_title_paragraphs(mission, profile):
    ent = normalize_text(mission.get('entreprise'))
    poste = normalize_text(mission.get('poste')).replace(' et ', ' & ')
    duree = display_mission_duration(mission.get('duree'))
    segments = [(kind, text) for kind, text in [('ent', ent), ('poste', poste), ('duree', duree)] if text]
    if not segments:
        return [[('ent', '')]]

    def joined(parts):
        row = []
        for idx, (kind, text) in enumerate(parts):
            if idx:
                row.append(('sep', f' {EN} '))
            row.append((kind, text))
        return row

    full = ' '.join(
        f'{text}' if idx == 0 else f'{EN} {text}'
        for idx, (_, text) in enumerate(segments)
    ).strip()
    if estimate_lines(full, profile['title_char_limit']) <= 2:
        return [joined(segments)]

    if len(segments) == 3:
        second = f"{poste} {EN} {duree}".strip()
        if estimate_lines(ent, profile['title_char_limit']) <= 2 and estimate_lines(second, profile['title_char_limit']) <= 2:
            return [
                [('ent', ent)],
                [('poste', poste), ('sep', f' {EN} '), ('duree', duree)],
            ]

    if len(segments) == 2:
        first_kind, first_text = segments[0]
        second_kind, second_text = segments[1]
        return [
            [(first_kind, first_text)],
            [(second_kind, second_text)],
        ]

    if estimate_lines(ent, profile['title_char_limit']) <= 2 and len(segments) == 3:
        return [
            [('ent', ent)],
            [('poste', poste), ('sep', f' {EN} '), ('duree', duree)],
        ]

    return [[(kind, text)] for kind, text in segments]

def mission_title_line_count(mission, profile):
    total = 0
    for paragraph in mission_title_paragraphs(mission, profile):
        text = ''.join(text for _, text in paragraph)
        total += estimate_lines(text, profile['title_char_limit'])
    return max(total, 1)

def title_profile(profile, mission):
    total_lines = mission_title_line_count(mission, profile)
    tuned = dict(profile)
    if total_lines >= 3:
        tuned['title_font_sz'] = max(profile['title_font_sz'] - 125, 1200)
        tuned['sep_font_sz'] = max(profile['sep_font_sz'] - 125, 1200)
        tuned['title_kern'] = max(profile['title_kern'] - 150, 600)
    return tuned

def mission_title_lines(mission):
    return mission_title_lines_for_profile(mission, LAYOUT_PROFILES[0])

def mission_zone_h(realisations):
    return mission_zone_h_for_profile(realisations, LAYOUT_PROFILES[0])

def mission_title_lines_for_profile(mission, profile):
    return mission_title_line_count(mission, profile)

def mission_zone_h_for_profile(realisations, profile):
    lines = sum(estimate_lines(real, profile['bullet_char_limit']) for real in realisations[:MAX_REALISATIONS_PER_MISSION]) or 1
    return lines * profile['bullet_h'] + profile['bullet_overhead']

def mission_layout(mission, profile=None):
    profile = profile or LAYOUT_PROFILES[0]
    tuned = title_profile(profile, mission)
    extra_lines = max(mission_title_lines_for_profile(mission, tuned) - 1, 0)
    extra_h = extra_lines * tuned['title_line_h']
    return {
        'group_h': tuned['group_h'] + extra_h,
        'zone_offset': ZONE1_TOP_OFFSET + extra_h,
        'zone_h': mission_zone_h_for_profile(mission.get('realisations', []), tuned),
    }

def chunk_bottom(chunk, profile=None):
    profile = profile or LAYOUT_PROFILES[0]
    if not chunk:
        return ZONE1_Y + mission_zone_h_for_profile([], profile)
    current_bottom = None
    for idx, mission in enumerate(chunk[:3]):
        layout = mission_layout(mission, profile)
        grp_y = GRP1_Y if idx == 0 else current_bottom + GAP_GRP
        zone_y = grp_y + layout['zone_offset']
        current_bottom = zone_y + layout['zone_h']
    return current_bottom

def chunk_gap(chunk, profile=None):
    """Gap entre missions. Pour 2 missions, gap plus généreux pour réduire le blanc en bas."""
    profile = profile or LAYOUT_PROFILES[0]
    if len(chunk) <= 1:
        return 90000
    base_gap = TEMPLATE_MISSION_GAP
    if profile['name'] == 'dense':
        base_gap = 70000
    elif profile['name'] == 'compact':
        base_gap = 85000
    elif profile['name'] == 'airy':
        base_gap = 95000
    if len(chunk) == 2:
        # Espace libre sous le contenu
        content_h = chunk_bottom(chunk, profile) - GRP1_Y
        available_h = SLIDE_CONTENT_LIMIT - GRP1_Y
        free = max(available_h - content_h, 0)
        # Absorber ~20% du blanc dans le gap — le reste en bas
        extra = min(free // 5, 400000)
        return base_gap + extra
    return base_gap

def chunk_shift(chunk, profile=None):
    """Pas de décalage : le contenu commence toujours en haut."""
    return 0

# ── RPR ───────────────────────────────────────────────────────────────────────
def rpr_blanc(sz, bold=False, italic=False):
    r = etree.Element(f'{A}rPr')
    r.set('lang','fr-FR'); r.set('sz',str(sz)); r.set('dirty','0')
    if bold:   r.set('b','1')
    if italic: r.set('i','1')
    etree.SubElement(etree.SubElement(r,f'{A}solidFill'),f'{A}schemeClr').set('val','bg1')
    etree.SubElement(r,f'{A}latin').set('typeface','+mn-lt')
    return r

def rpr_cyan(sz, bold=True, italic=False):
    r = etree.Element(f'{A}rPr')
    r.set('lang','fr-FR'); r.set('sz',str(sz)); r.set('dirty','0')
    if bold:   r.set('b','1')
    if italic: r.set('i','1')
    etree.SubElement(etree.SubElement(r,f'{A}solidFill'),f'{A}srgbClr').set('val',CY)
    etree.SubElement(r,f'{A}latin').set('typeface','+mn-lt')
    return r

def rpr_bullet(profile=None):
    profile = profile or LAYOUT_PROFILES[0]
    r = etree.Element(f'{A}rPr')
    r.set('lang','fr-FR'); r.set('sz',str(profile['bullet_font_sz'])); r.set('dirty','0')
    clr = etree.SubElement(etree.SubElement(r,f'{A}solidFill'),f'{A}prstClr')
    clr.set('val', 'black')
    etree.SubElement(clr,f'{A}lumMod').set('val','75000')
    etree.SubElement(clr,f'{A}lumOff').set('val','25000')
    return r

def rpr_titre_mission(profile=None):
    profile = profile or LAYOUT_PROFILES[0]
    r = etree.Element(f'{A}rPr')
    r.set('lang','fr-FR'); r.set('sz',str(profile['title_font_sz'])); r.set('b','1'); r.set('dirty','0')
    etree.SubElement(etree.SubElement(r,f'{A}solidFill'),f'{A}srgbClr').set('val', STYLE_MEMORY['mission_title_color'])
    lat = etree.SubElement(r,f'{A}latin')
    lat.set('typeface', STYLE_MEMORY['mission_title_font']); lat.set('panose','00000500000000000000')
    lat.set('pitchFamily','2'); lat.set('charset','0')
    return r

def rpr_sep(profile=None):
    profile = profile or LAYOUT_PROFILES[0]
    r = etree.Element(f'{A}rPr')
    for k,v in [('kumimoji','0'),('lang','fr-FR'),('sz',str(profile['sep_font_sz'])),('b','1'),
                ('u','none'),('strike','noStrike'),('kern',str(profile['title_kern'])),('dirty','0')]:
        r.set(k,v)
    etree.SubElement(etree.SubElement(r,f'{A}ln'),f'{A}noFill')
    etree.SubElement(etree.SubElement(r,f'{A}solidFill'),f'{A}srgbClr').set('val', STYLE_MEMORY['mission_title_color'])
    etree.SubElement(r,f'{A}effectLst')
    etree.SubElement(r,f'{A}uLnTx'); etree.SubElement(r,f'{A}uFillTx')
    lat = etree.SubElement(r,f'{A}latin')
    lat.set('typeface', STYLE_MEMORY['mission_title_font']); lat.set('panose','00000500000000000000')
    lat.set('pitchFamily','2'); lat.set('charset','0')
    return r

def mk_run(rpr, text):
    r = etree.Element(f'{A}r')
    r.append(copy.deepcopy(rpr))
    etree.SubElement(r,f'{A}t').text = text
    return r

def clear_p(tb):
    for p in list(tb.findall(f'{A}p')): tb.remove(p)

def clear_text_body(tb):
    if tb is None:
        return
    clear_p(tb)
    p = etree.SubElement(tb, f'{A}p')
    end = etree.SubElement(p, f'{A}endParaRPr')
    end.set('lang', 'fr-FR')

def get_name(el):
    for tag in [f'{P}nvSpPr',f'{P}nvGrpSpPr']:
        nv = el.find(tag)
        if nv is not None:
            c = nv.find(f'{P}cNvPr')
            if c is not None: return c.get('name','')
    return ''

def get_tb(el):
    tb = el.find(f'{P}txBody')
    return tb if tb is not None else el.find(f'{A}txBody')

def get_xfrm(el):
    for sptag in [f'{P}spPr', f'{P}grpSpPr']:
        sp = el.find(sptag)
        if sp is not None:
            xfrm = sp.find(f'{A}xfrm')
            if xfrm is not None:
                return xfrm, sp
    return None, None

def set_shape_frame(el, frame):
    """Reset off/ext to template values. Preserve chOff/chExt for groups
    (they define the child coordinate space and must not equal off)."""
    if el is None or not frame:
        return
    xfrm, _ = get_xfrm(el)
    if xfrm is None:
        return
    off = xfrm.find(f'{A}off')
    ext = xfrm.find(f'{A}ext')
    if off is not None:
        off.set('x', str(int(frame['x'])))
        off.set('y', str(int(frame['y'])))
    if ext is not None:
        ext.set('cx', str(int(frame['cx'])))
        ext.set('cy', str(int(frame['cy'])))

def reset_shape_to_template(el, name):
    frame = TEMPLATE_MEMORY.get(name)
    if frame:
        set_shape_frame(el, frame)
    if name in EDITABLE_TEXT_SHAPES:
        clear_text_body(get_tb(el))

def reset_slide_to_template(root, shapes):
    for name in TEXT_SHAPES:
        el = shapes.get(name)
        if el is not None:
            reset_shape_to_template(el, name)
    for name in ['Groupe 20', 'ZoneTexte 22', 'ZoneTexte 23']:
        el = shapes.get(name)
        if el is None:
            continue
        clear_text_body(get_tb(el))
        xfrm, _ = get_xfrm(el)
        if xfrm is None:
            continue
        off = xfrm.find(f'{A}off')
        if off is not None:
            off.set('y', '99999999')

# ── PARAGRAPHES ───────────────────────────────────────────────────────────────
def para_bullet_mission(text, profile=None):
    profile = profile or LAYOUT_PROFILES[0]
    p = etree.Element(f'{A}p')
    pPr = etree.SubElement(p,f'{A}pPr')
    pPr.set('marL','1085850'); pPr.set('lvl','2'); pPr.set('indent','-171450')
    etree.SubElement(etree.SubElement(pPr,f'{A}lnSpc'),f'{A}spcPct').set('val',str(profile['bullet_line_spc']))
    etree.SubElement(etree.SubElement(pPr,f'{A}spcAft'),f'{A}spcPts').set('val',str(profile['bullet_spc_aft']))
    bf = etree.SubElement(pPr,f'{A}buFont')
    bf.set('typeface','Arial'); bf.set('panose','020B0604020202020204')
    bf.set('pitchFamily','34'); bf.set('charset','0')
    etree.SubElement(pPr,f'{A}buChar').set('char','•')
    tab = etree.SubElement(pPr,f'{A}tabLst')
    t = etree.SubElement(tab,f'{A}tab'); t.set('pos','457200'); t.set('algn','l')
    etree.SubElement(pPr,f'{A}defRPr')
    p.append(mk_run(rpr_bullet(profile), text))
    return p

def para_bullet_col(text, marL='257175', indent='-257175'):
    p = etree.Element(f'{A}p')
    pPr = etree.SubElement(p,f'{A}pPr')
    pPr.set('marL',marL); pPr.set('indent',indent)
    bf = etree.SubElement(pPr,f'{A}buFont')
    bf.set('typeface','Arial'); bf.set('panose','020B0604020202020204')
    bf.set('pitchFamily','34'); bf.set('charset','0')
    etree.SubElement(pPr,f'{A}buChar').set('char','•')
    p.append(mk_run(rpr_blanc(1300), text))
    return p

def para_comp():
    pPr = etree.Element(f'{A}pPr')
    pPr.set('marL','0'); pPr.set('indent','0')
    etree.SubElement(etree.SubElement(pPr,f'{A}lnSpc'),f'{A}spcPct').set('val','95000')
    etree.SubElement(etree.SubElement(pPr,f'{A}spcBef'),f'{A}spcPts').set('val','0')
    etree.SubElement(etree.SubElement(pPr,f'{A}spcAft'),f'{A}spcPts').set('val','0')
    etree.SubElement(pPr,f'{A}buNone')
    return pPr

# ── SHAPE UPDATERS ────────────────────────────────────────────────────────────
def titre_sz(n):
    if n<=20: return '2700'
    elif n<=30: return '2400'
    elif n<=40: return '2000'
    else: return '1800'

def cyan_sz(n):
    if n<=45: return '1900'
    elif n<=58: return '1600'
    else: return '1400'

def upd_titre(el, prenom, n, total, titre, xp):
    tb = get_tb(el)
    if tb is None: return
    clear_p(tb)
    l2 = f'{titre} ({n}/{total})'
    sz = titre_sz(len(l2))
    # Ligne 1 : prénom
    p1 = etree.SubElement(tb,f'{A}p')
    etree.SubElement(p1,f'{A}pPr').set('algn','l')
    p1.append(mk_run(rpr_blanc(sz), prenom))
    # Ligne 2 : titre (n/total)
    p2 = etree.SubElement(tb,f'{A}p')
    etree.SubElement(p2,f'{A}pPr').set('algn','l')
    p2.append(mk_run(rpr_blanc(sz), l2))
    ep = etree.SubElement(p2,f'{A}endParaRPr')
    ep.set('lang','fr-FR'); ep.set('sz','5400'); ep.set('dirty','0')
    etree.SubElement(etree.SubElement(ep,f'{A}solidFill'),f'{A}schemeClr').set('val','bg1')
    etree.SubElement(ep,f'{A}latin').set('typeface','+mn-lt')
    # Ligne 3 : X ans cyan italique
    l3 = f"{xp} d'expériences en {titre}".strip()
    p3 = etree.SubElement(tb,f'{A}p')
    etree.SubElement(p3,f'{A}pPr').set('algn','l')
    p3.append(mk_run(rpr_cyan(cyan_sz(len(l3)), bold=False, italic=True), l3))

def upd_competences(el, comps):
    tb = get_tb(el)
    if tb is None: return
    clear_p(tb)
    mid = len(comps)//2
    l1 = ' '.join(comps[:mid]) if mid>0 else ' '.join(comps)
    l2 = ' '.join(comps[mid:]) if mid>0 and mid<len(comps) else None
    for l in filter(None,[l1,l2]):
        p = etree.SubElement(tb,f'{A}p')
        p.append(para_comp())
        p.append(mk_run(rpr_cyan(1600), l))
    p3 = etree.SubElement(tb,f'{A}p')
    p3.append(para_comp())
    ep = etree.SubElement(p3,f'{A}endParaRPr')
    ep.set('lang','fr-FR'); ep.set('sz','1350'); ep.set('b','1'); ep.set('dirty','0')
    etree.SubElement(etree.SubElement(ep,f'{A}solidFill'),f'{A}srgbClr').set('val',CY)

def upd_col(el, label, items, max_i, marL='257175', indent='-257175'):
    tb = get_tb(el)
    if tb is None: return
    clear_p(tb)
    # Label
    p = etree.SubElement(tb,f'{A}p')
    p.append(mk_run(rpr_blanc(1800,bold=True), label))
    # Ligne vide après label
    pv = etree.SubElement(tb,f'{A}p')
    ep = etree.SubElement(pv,f'{A}endParaRPr')
    ep.set('lang','fr-FR'); ep.set('sz','1800'); ep.set('b','1'); ep.set('dirty','0')
    etree.SubElement(etree.SubElement(ep,f'{A}solidFill'),f'{A}schemeClr').set('val','bg1')
    for item in items[:max_i]:
        tb.append(para_bullet_col(item, marL, indent))

def upd_aptitudes(el, items):
    tb = get_tb(el)
    if tb is None: return
    clear_p(tb)
    # Ligne vide initiale (template a ça)
    pv0 = etree.SubElement(tb,f'{A}p')
    ep0 = etree.SubElement(pv0,f'{A}endParaRPr')
    ep0.set('lang','fr-FR'); ep0.set('sz','1800'); ep0.set('b','1'); ep0.set('dirty','0')
    etree.SubElement(etree.SubElement(ep0,f'{A}solidFill'),f'{A}schemeClr').set('val','bg1')
    p = etree.SubElement(tb,f'{A}p')
    p.append(mk_run(rpr_blanc(1800,bold=True), 'APTITUDES'))
    pv = etree.SubElement(tb,f'{A}p')
    ep = etree.SubElement(pv,f'{A}endParaRPr')
    ep.set('lang','fr-FR'); ep.set('sz','1350'); ep.set('dirty','0')
    etree.SubElement(etree.SubElement(ep,f'{A}solidFill'),f'{A}schemeClr').set('val','bg1')
    for item in items[:5]: tb.append(para_bullet_col(item))

def upd_grp_titre(grp, ent, poste, duree, profile=None):
    """Met à jour le texte du titre de mission dans le groupe."""
    profile = profile or LAYOUT_PROFILES[0]
    mission = {'entreprise': ent, 'poste': poste, 'duree': duree}
    tuned = title_profile(profile, mission)
    paragraphs = mission_title_paragraphs(mission, tuned)
    for sp in grp.iter(f'{P}sp'):
        tb = get_tb(sp)
        if tb is None: continue
        clear_p(tb)
        for idx, parts in enumerate(paragraphs):
            p = etree.SubElement(tb,f'{A}p')
            pPr = etree.SubElement(p,f'{A}pPr')
            for k,v in [('marL','0'),('marR','0'),('lvl','0'),('indent','0'),('algn','l'),
                        ('defTabSz','914400'),('rtl','0'),('eaLnBrk','1'),('fontAlgn','auto'),
                        ('latinLnBrk','0'),('hangingPunct','1')]:
                pPr.set(k,v)
            etree.SubElement(etree.SubElement(pPr,f'{A}lnSpc'),f'{A}spcPct').set('val','100000')
            etree.SubElement(etree.SubElement(pPr,f'{A}spcBef'),f'{A}spcPts').set('val','300' if idx == 0 else '0')
            etree.SubElement(etree.SubElement(pPr,f'{A}spcAft'),f'{A}spcPts').set('val','0')
            etree.SubElement(pPr,f'{A}buClrTx'); etree.SubElement(pPr,f'{A}buSzTx')
            bf = etree.SubElement(pPr,f'{A}buFont')
            bf.set('typeface','Arial'); bf.set('pitchFamily','34'); bf.set('charset','0')
            etree.SubElement(pPr,f'{A}buNone')
            etree.SubElement(pPr,f'{A}tabLst'); etree.SubElement(pPr,f'{A}defRPr')
            for part_type, text in parts:
                if not text:
                    continue
                if part_type == 'sep':
                    p.append(mk_run(rpr_sep(tuned), text))
                else:
                    p.append(mk_run(rpr_titre_mission(tuned), text))
        break

def upd_zone_bullets(el, reals, profile=None):
    """Met à jour les bullets d'une zone mission."""
    profile = profile or LAYOUT_PROFILES[0]
    tb = get_tb(el)
    if tb is None: return
    clear_p(tb)
    reals = reals[:MAX_REALISATIONS_PER_MISSION]
    for r in reals: tb.append(para_bullet_mission(r, profile))
    # Paragraphe vide final
    p_end = etree.SubElement(tb,f'{A}p')
    pPr = etree.SubElement(p_end,f'{A}pPr')
    pPr.set('marL','1085850'); pPr.set('lvl','2'); pPr.set('indent','-171450')
    etree.SubElement(etree.SubElement(pPr,f'{A}spcBef'),f'{A}spcPts').set('val','0')
    etree.SubElement(pPr,f'{A}buChar').set('char','•')
    etree.SubElement(p_end,f'{A}endParaRPr').set('lang','fr-FR')

# ── REPOSITIONNEMENT ──────────────────────────────────────────────────────────
def reposition(root, chunk, profile=None):
    """
    Repositionne les groupes/zones missions dynamiquement.
    Conserve delta chOff-off pour que les enfants des groupes suivent.
    Masque les zones non utilisées hors slide.
    """
    # Collecter toutes les shapes
    shapes = {}
    for el in root.iter():
        for tag in [f'{P}nvSpPr',f'{P}nvGrpSpPr']:
            nv = el.find(tag)
            if nv is not None:
                c = nv.find(f'{P}cNvPr')
                if c is not None and c.get('name'):
                    shapes[c.get('name')] = el

    def move_y(name, new_y):
        el = shapes.get(name)
        if el is None: return
        # Groupe : only change off, keep chOff fixed (it defines child coordinate space)
        grpSpPr = el.find(f'{P}grpSpPr')
        if grpSpPr is not None:
            xfrm = grpSpPr.find(f'{A}xfrm')
            if xfrm is not None:
                off = xfrm.find(f'{A}off')
                if off is not None:
                    off.set('y', str(int(new_y)))
            return
        # Shape simple : spPr
        spPr = el.find(f'{P}spPr')
        if spPr is not None:
            xfrm = spPr.find(f'{A}xfrm')
            if xfrm is not None:
                off = xfrm.find(f'{A}off')
                if off is not None: off.set('y', str(int(new_y)))

    def clear_text(name):
        el = shapes.get(name)
        if el is None:
            return
        tb = get_tb(el)
        if tb is not None:
            clear_text_body(tb)

    def resize_h(name, new_h):
        el = shapes.get(name)
        if el is None:
            return
        grpSpPr = el.find(f'{P}grpSpPr')
        if grpSpPr is not None:
            xfrm = grpSpPr.find(f'{A}xfrm')
            if xfrm is not None:
                ext = xfrm.find(f'{A}ext')
                if ext is not None:
                    ext.set('cy', str(int(new_h)))
            return
        spPr = el.find(f'{P}spPr')
        if spPr is not None:
            xfrm = spPr.find(f'{A}xfrm')
            if xfrm is not None:
                ext = xfrm.find(f'{A}ext')
                if ext is not None:
                    ext.set('cy', str(int(new_h)))

    PAIRS = [
        ('Groupe 6',  MISSION_TITLE_NAMES[0], MISSION_ZONE_NAMES[0]),
        ('Groupe 9',  MISSION_TITLE_NAMES[1], MISSION_ZONE_NAMES[1]),
        ('Groupe 11', MISSION_TITLE_NAMES[2], MISSION_ZONE_NAMES[2]),
    ]

    current_bottom = None
    profile = profile or LAYOUT_PROFILES[0]
    gap = chunk_gap(chunk, profile)
    shift = chunk_shift(chunk, profile)

    for i in range(len(PAIRS)):
        grp_name, title_name, zone_name = PAIRS[i]
        if i < len(chunk):
            layout = mission_layout(chunk[i], profile)
            new_grp_y = (GRP1_Y + shift) if i == 0 else current_bottom + gap
            new_zone_y = new_grp_y + layout['zone_offset']
            move_y(grp_name,  new_grp_y)
            move_y(zone_name, new_zone_y)
            resize_h(grp_name, layout['group_h'])
            resize_h(zone_name, layout['zone_h'])
            current_bottom = new_zone_y + layout['zone_h']
        else:
            # Masquer hors slide
            clear_text(title_name)
            clear_text(zone_name)
            move_y(grp_name,  99999999)
            move_y(zone_name, 99999999)

    # Toujours masquer les shapes fantômes du template contaminé
    for name in ['Groupe 20','ZoneTexte 22','ZoneTexte 23']:
        clear_text(name)
        move_y(name, 99999999)

# ── ROLES ─────────────────────────────────────────────────────────────────────
ROLES = {
    'Titre 1':                    'titre',
    'Espace réservé du texte 5':  'comp',
    'Espace réservé du texte 13': 'skip',   # Label "COMPÉTENCES CLÉS" — ne pas toucher
    'Espace réservé du texte 2':  'formation',
    'Espace réservé du texte 3':  'outils',
    'Espace réservé du texte 4':  'aptitudes',
    'Image 8':                    'skip',   # Logo V4F — ne pas toucher
    'Rectangle 7':                'skip',   # Fond bleu — ne pas toucher
    'Groupe 6':   'skip', 'ZoneTexte 14': 'g0', 'ZoneTexte 15': 'z0',
    'Groupe 9':   'skip', 'ZoneTexte 16': 'g1', 'ZoneTexte 18': 'z1',
    'Groupe 11':  'skip', 'ZoneTexte 17': 'g2', 'ZoneTexte 19': 'z2',
    'Groupe 20':  'skip', 'ZoneTexte 22': 'skip', 'ZoneTexte 23': 'skip',
}

def process_slide(xml, data, slide_n, total, chunk, profile=None):
    root = etree.fromstring(xml)
    d = data
    profile = profile or LAYOUT_PROFILES[0]

    # Build shapes dict (last-wins for duplicates = group child)
    shapes = {}
    groups = {}
    for el in root.iter():
        name = get_name(el)
        if name:
            shapes[name] = el
            if el.tag == f'{P}grpSp':
                groups[name] = el

    # Hide standalone shapes that duplicate group children (they retain placeholder text)
    spTree = root.find(f'{P}cSld/{P}spTree')
    if spTree is not None:
        for sp in list(spTree.findall(f'{P}sp')):
            name = get_name(sp)
            if name in STANDALONE_DUPLICATES:
                clear_text_body(get_tb(sp))
                xfrm_el, _ = get_xfrm(sp)
                if xfrm_el is not None:
                    off = xfrm_el.find(f'{A}off')
                    if off is not None:
                        off.set('y', '99999999')

    reset_slide_to_template(root, shapes)

    def apply(name, fn, *args):
        el = shapes.get(name)
        if el is not None:
            fn(el, *args)

    apply('Titre 1', upd_titre, d['prenom'], slide_n, total, d['titre'], d['annees_xp'])
    apply('Espace réservé du texte 5', upd_competences, d['competences_cles'])
    apply('Espace réservé du texte 2', upd_col, 'FORMATION', d['formation'], 5)
    apply('Espace réservé du texte 3', upd_col, 'OUTILS MAITRISES', d['outils'], 5, '285750', '-285750')
    apply('Espace réservé du texte 4', upd_aptitudes, d['aptitudes'])

    for idx, mission in enumerate(chunk[:3]):
        # Use GROUP element for title (not standalone shape which is now hidden)
        grp = groups.get(GROUP_NAMES[idx])
        if grp is not None:
            upd_grp_titre(grp, mission['entreprise'], mission['poste'], mission['duree'], profile)
        apply(MISSION_ZONE_NAMES[idx], upd_zone_bullets, mission.get('realisations', []), profile)

    reposition(root, chunk, profile)
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

# ── DÉCOUPAGE ADAPTATIF ───────────────────────────────────────────────────────
def best_profile_for_chunk(chunk):
    if len(chunk) > MAX_MISSIONS_PER_SLIDE:
        return None
    fitting_profiles = [
        profile for profile in LAYOUT_PROFILES
        if chunk_bottom(chunk, profile) <= SLIDE_CONTENT_LIMIT
    ]
    if not fitting_profiles:
        return None
    if len(chunk) >= 3:
        comfortable = [
            profile for profile in fitting_profiles
            if (chunk_bottom(chunk, profile) / SLIDE_CONTENT_LIMIT) <= THREE_MISSION_MAX_FILL
        ]
        if comfortable:
            return comfortable[0]
        return None
    return fitting_profiles[0]

def compact_underfilled_slide(chunk, profile):
    if not chunk:
        return profile
    fill = chunk_bottom(chunk, profile) / SLIDE_CONTENT_LIMIT
    if fill >= 0.8:
        return profile
    if len(chunk) == 1:
        return profile
    # Pour les slides sous-remplies, on préfère le profil qui occupe le plus
    # d'espace vertical sans déborder, au lieu d'alléger encore la slide.
    fitting = [
        candidate for candidate in LAYOUT_PROFILES
        if chunk_bottom(chunk, candidate) <= SLIDE_CONTENT_LIMIT
    ]
    if not fitting:
        return profile
    return max(fitting, key=lambda candidate: chunk_bottom(chunk, candidate))

def plan_slides(missions):
    """Cherche la meilleure répartition globale des missions et le profil de chaque slide."""
    missions = missions[:MAX_MISSIONS_TOTAL]
    if not missions:
        return [{'missions': [], 'profile': LAYOUT_PROFILES[0]}]

    profile_index = {profile['name']: idx for idx, profile in enumerate(LAYOUT_PROFILES)}
    profile_by_name = {profile['name']: profile for profile in LAYOUT_PROFILES}

    def score(spec_plan):
        fills = [
            chunk_bottom(missions[start:end], profile_by_name[profile_name]) / SLIDE_CONTENT_LIMIT
            for start, end, profile_name in spec_plan
        ]
        whitespace = [1 - min(fill, 1) for fill in fills]
        dense_penalty = sum(profile_index[profile_name] for _, _, profile_name in spec_plan)
        single_mission_penalty = sum(1 for start, end, _ in spec_plan if (end - start) == 1)
        frontload_priority = tuple(-(end - start) for start, end, _ in spec_plan)
        return (
            len(spec_plan),
            single_mission_penalty,
            frontload_priority,
            round(sum(w * w for w in whitespace), 6),
            round(sum(whitespace), 6),
            round(max(fills) - min(fills), 6),
            dense_penalty,
        )

    @lru_cache(maxsize=None)
    def solve(start):
        if start >= len(missions):
            return tuple()

        best_score = None
        best_plan = None

        max_size = min(MAX_MISSIONS_PER_SLIDE, len(missions) - start)
        for size in range(max_size, 0, -1):
            end = start + size
            chunk = missions[start:end]
            profile = best_profile_for_chunk(chunk)
            if profile is None:
                continue
            profile = compact_underfilled_slide(chunk, profile)
            suffix = solve(end)
            if suffix is None:
                continue
            candidate = ((start, end, profile['name']),) + suffix
            candidate_score = score(candidate)
            if best_score is None or candidate_score < best_score:
                best_score = candidate_score
                best_plan = candidate
        return best_plan

    best = solve(0)
    if best is not None:
        return [
            {'missions': missions[start:end], 'profile': profile_by_name[profile_name]}
            for start, end, profile_name in best
        ]

    # Fallback: 1 mission par slide avec le profil le plus compact disponible.
    fallback_profile = LAYOUT_PROFILES[-1]
    return [{'missions': [m], 'profile': fallback_profile} for m in missions]

# ── BUILD PPTX ────────────────────────────────────────────────────────────────
def build_pptx(data, output_path):
    slide_plan = plan_slides(data.get('missions',[]))
    total  = len(slide_plan)

    with zipfile.ZipFile(str(TEMPLATE_PATH),'r') as z:
        tpl = {name: z.read(name) for name in z.namelist()}

    slide1_xml = tpl['ppt/slides/slide1.xml']
    prs   = etree.fromstring(tpl['ppt/presentation.xml'])
    prels = etree.fromstring(tpl.get('ppt/_rels/presentation.xml.rels', b'<Relationships/>'))
    sNS   = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
    sldIdLst = prs.find(f'.//{sNS}sldIdLst')

    maxr = max((int(r.get('Id','rId0').replace('rId','')) for r in prels.findall(f'{{{RN}}}Relationship')), default=10)
    maxs = max((int(s.get('id','255')) for s in (sldIdLst.findall(f'{sNS}sldId') if sldIdLst is not None else [])), default=255)

    out = dict(tpl)
    rels1 = tpl.get('ppt/slides/_rels/slide1.xml.rels', b'')
    out['ppt/slides/slide1.xml'] = process_slide(slide1_xml, data, 1, total, slide_plan[0]['missions'], slide_plan[0]['profile'])

    for i, item in enumerate(slide_plan[1:], start=2):
        out[f'ppt/slides/slide{i}.xml'] = process_slide(slide1_xml, data, i, total, item['missions'], item['profile'])
        if rels1:
            out[f'ppt/slides/_rels/slide{i}.xml.rels'] = rels1
        ct = etree.fromstring(out['[Content_Types].xml'])
        if f'slide{i}.xml' not in out['[Content_Types].xml'].decode():
            ov = etree.SubElement(ct,'Override')
            ov.set('PartName', f'/ppt/slides/slide{i}.xml')
            ov.set('ContentType','application/vnd.openxmlformats-officedocument.presentationml.slide+xml')
            out['[Content_Types].xml'] = etree.tostring(ct, xml_declaration=True, encoding='UTF-8', standalone=True)
        maxr += 1
        rId = f'rId{maxr}'
        nr = etree.SubElement(prels, f'{{{RN}}}Relationship')
        nr.set('Id',rId)
        nr.set('Type','http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide')
        nr.set('Target',f'slides/slide{i}.xml')
        if sldIdLst is not None:
            maxs += 1
            ns = etree.SubElement(sldIdLst, f'{sNS}sldId')
            ns.set('id', str(maxs))
            ns.set(f'{R}id', rId)

    out['ppt/presentation.xml'] = etree.tostring(prs, xml_declaration=True, encoding='UTF-8', standalone=True)
    out['ppt/_rels/presentation.xml.rels'] = etree.tostring(prels, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(str(output_path),'w',zipfile.ZIP_DEFLATED) as z:
        for name, content in out.items():
            z.writestr(name, content)

# ── LLM ──────────────────────────────────────────────────────────────────────
PROMPT_EXTRACT = """Tu extrais un CV pour alimenter un template PowerPoint strict. Retourne UNIQUEMENT un JSON valide.

{
  "prenom": "Prénom",
  "nom": "Nom",
  "titre": "Titre court 3-5 mots",
  "annees_xp": "X ans",
  "competences_cles": ["max 4, 2-4 mots chacune"],
  "formation": ["max 3, format: Diplôme - École"],
  "outils": ["max 5, grouper avec /"],
  "aptitudes": ["max 4, courtes"],
  "missions": [
    {"entreprise":"","poste":"","duree":"dates ou durée du CV","realisations":["verbe + résultat, formulation du CV conservée si possible"]}
  ]
}

RÈGLES IMPÉRATIVES:
- N'invente rien. Utilise uniquement des informations explicitement présentes dans le CV.
- Si une information n'est pas clairement visible dans le CV, mets "" ou [].
- Reprends autant que possible les formulations exactes du CV, surtout pour entreprise, poste, diplôme, outils.
- Ne déduis pas d'années d'expérience à partir des dates si ce n'est pas explicitement écrit.
- Le champ "titre" doit être repris du poste le plus récent ou d'un intitulé clairement écrit dans le CV, sans reformulation marketing.
- Les "competences_cles", "outils" et "aptitudes" doivent être des éléments réellement visibles dans le CV, pas des extrapolations.
- Les "realisations" doivent garder le wording du CV autant que possible. Raccourcis seulement si c'est vraiment nécessaire pour l'affichage.
- Garde toutes les expériences visibles et significatives du CV, en chronologie inversée.
- Max 24 missions, max 5 réalisations par mission."""

PROMPT_VALID = """Tu reçois un CV source et un JSON extrait. Corrige le JSON en restant strictement fidèle au CV.

RÈGLES:
- Supprime toute information absente, douteuse ou reformulée de manière trop libre.
- Garde uniquement ce qui est ancré dans le CV.
- Préfère une valeur vide plutôt qu'une invention.
- Respecte le format PPT: titre court, max 4 compétences, max 3 formations, max 5 outils, max 4 aptitudes, max 24 missions, max 5 réalisations par mission.
- Missions en chronologie inversée.

Retourne UNIQUEMENT le JSON corrigé.

CV:
{cv_text}

JSON:
{json_data}"""

PROMPT_REPAIR_JSON = """Tu reçois un texte censé être un JSON pour un CV, mais il peut être mal formé, tronqué ou entouré de texte parasite.

Répare-le en JSON valide strict.

RÈGLES:
- Retourne UNIQUEMENT un JSON valide.
- N'ajoute aucune information qui n'est pas déjà présente dans le texte reçu.
- Si un champ est incertain ou incomplet, mets une chaîne vide "" ou une liste [].
- Conserve la structure attendue:
{{
  "prenom": "",
  "nom": "",
  "titre": "",
  "annees_xp": "",
  "competences_cles": [],
  "formation": [],
  "outils": [],
  "aptitudes": [],
  "missions": [
    {{"entreprise":"","poste":"","duree":"","realisations":[]}}
  ]
}}

Texte à réparer:
{raw_text}"""

def call_llm(messages, temp=0.1, json_mode=False):
    client = ollama.Client(host=OLLAMA_HOST)
    kwargs = {
        'model': MODEL,
        'messages': messages,
        'options': {'temperature':temp,'num_ctx':8192},
    }
    if json_mode:
        kwargs['format'] = 'json'
    resp = client.chat(**kwargs)
    raw = resp['message']['content'].strip()
    raw = re.sub(r'```json\s*','',raw); raw = re.sub(r'```\s*','',raw)
    return raw.strip()

def extract_balanced_json(raw):
    raw = str(raw or '')
    start = None
    for idx, ch in enumerate(raw):
        if ch in '{[':
            start = idx
            break
    if start is None:
        return None
    pairs = {'{': '}', '[': ']'}
    stack = []
    in_string = False
    escape = False
    for idx in range(start, len(raw)):
        ch = raw[idx]
        if in_string:
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
            continue
        if ch in '{[':
            stack.append(ch)
            continue
        if ch in '}]':
            if not stack:
                continue
            opener = stack[-1]
            if pairs[opener] == ch:
                stack.pop()
                if not stack:
                    return raw[start:idx + 1]
    return None

def auto_close_json(raw):
    raw = str(raw or '').strip()
    if not raw:
        return raw
    start = None
    for idx, ch in enumerate(raw):
        if ch in '{[':
            start = idx
            break
    if start is not None:
        raw = raw[start:]
    pairs = {'{': '}', '[': ']'}
    stack = []
    out = []
    in_string = False
    escape = False
    for ch in raw:
        out.append(ch)
        if in_string:
            if escape:
                escape = False
            elif ch == '\\':
                escape = True
            elif ch == '"':
                in_string = False
            continue
        if ch == '"':
            in_string = True
            continue
        if ch in '{[':
            stack.append(ch)
        elif ch in '}]' and stack and pairs[stack[-1]] == ch:
            stack.pop()
    candidate = ''.join(out).rstrip()
    if in_string:
        candidate += '"'
    candidate = re.sub(r',\s*$', '', candidate)
    if re.search(r':\s*$', candidate):
        candidate += ' ""'
    for opener in reversed(stack):
        candidate += pairs[opener]
    candidate = re.sub(r',(\s*[}\]])', r'\1', candidate)
    return candidate

def try_parse_json(raw):
    candidates = []
    cleaned = str(raw or '').strip()
    if cleaned:
        candidates.append(cleaned)
        balanced = extract_balanced_json(cleaned)
        if balanced and balanced not in candidates:
            candidates.append(balanced)
        no_trailing_commas = re.sub(r',(\s*[}\]])', r'\1', cleaned)
        if no_trailing_commas not in candidates:
            candidates.append(no_trailing_commas)
        balanced_no_trailing = extract_balanced_json(no_trailing_commas)
        if balanced_no_trailing and balanced_no_trailing not in candidates:
            candidates.append(balanced_no_trailing)
        auto_closed = auto_close_json(cleaned)
        if auto_closed and auto_closed not in candidates:
            candidates.append(auto_closed)
    for candidate in candidates:
        try:
            return json.loads(candidate)
        except Exception:
            continue
    return None

def parse_json(raw, repair=False):
    parsed = try_parse_json(raw)
    if parsed is not None:
        return parsed
    if repair:
        try:
            repaired_raw = call_llm(
                [{'role':'user','content':PROMPT_REPAIR_JSON.format(raw_text=str(raw or '')[:12000])}],
                temp=0.0,
                json_mode=True,
            )
            parsed = try_parse_json(repaired_raw)
            if parsed is not None:
                return parsed
            raise ValueError(f"JSON invalide après réparation: {repaired_raw[:200]}")
        except Exception:
            pass
    raise ValueError(f"JSON invalide: {str(raw or '')[:200]}")

def clean(data, cv_text=''):
    def trunc(t,n): return ' '.join(str(t).split()[:n])
    def similar_key(text):
        tokens = re.findall(r'[a-z0-9]{3,}', ascii_fold(text))
        return ' '.join(tokens[:8])
    def uniq(items, limit):
        out, seen = [], set()
        for item in items:
            key = ascii_fold(item)
            near = similar_key(item)
            if not key or key in seen or near in seen:
                continue
            if any(key.startswith(existing) or existing.startswith(key) for existing in seen):
                continue
            seen.add(key)
            if near:
                seen.add(near)
            out.append(item)
            if len(out) >= limit:
                break
        return out

    src_text = normalize_text(cv_text)
    src_folded = ascii_fold(src_text)
    src_tok = source_tokens(src_text)

    data.setdefault('prenom','Consultant'); data.setdefault('nom','')
    data.setdefault('titre',''); data.setdefault('annees_xp','')

    prenom = trim_words(normalize_text(data.get('prenom')), 40)
    nom = trim_words(normalize_text(data.get('nom')), 60)
    titre = fit_title_field(trunc(normalize_text(data.get('titre')), 6), 50)
    annees_xp = normalize_text(data.get('annees_xp'))[:20]

    if src_text and prenom and not grounded(prenom, src_text, src_tok, src_folded, min_ratio=1.0, min_hits=1):
        prenom = 'Consultant'
    if src_text and nom and not grounded(nom, src_text, src_tok, src_folded, min_ratio=0.8, min_hits=1):
        nom = ''
    if src_text and titre and not grounded(titre, src_text, src_tok, src_folded, min_ratio=0.5, min_hits=1):
        titre = ''
    if src_text and annees_xp and not grounded(annees_xp, src_text, src_tok, src_folded, min_ratio=1.0, min_hits=1):
        annees_xp = ''

    data['prenom'] = prenom or 'Consultant'
    data['nom'] = nom
    data['annees_xp'] = annees_xp

    competences = [
        fit_title_field(normalize_text(c), 50)
        for c in data.get('competences_cles', [])
        if normalize_text(c) and (not src_text or grounded(c, src_text, src_tok, src_folded, min_ratio=0.5, min_hits=1))
    ]
    formation = [
        fit_formation(f)
        for f in data.get('formation', [])
        if normalize_text(f) and (not src_text or grounded(f, src_text, src_tok, src_folded, min_ratio=0.5, min_hits=2))
    ]
    outils = [
        fit_title_field(normalize_text(o), 60)
        for o in data.get('outils', [])
        if normalize_text(o) and (not src_text or grounded(o, src_text, src_tok, src_folded, min_ratio=0.5, min_hits=1))
    ]
    aptitudes = [
        fit_title_field(normalize_text(a), 50)
        for a in data.get('aptitudes', [])
        if normalize_text(a) and (not src_text or grounded(a, src_text, src_tok, src_folded, min_ratio=0.5, min_hits=1))
    ]

    data['competences_cles'] = uniq(competences, 4)
    data['formation'] = uniq(formation, 3)
    data['outils'] = uniq(outils, 5)
    data['aptitudes'] = uniq(aptitudes, 4)

    missions = []
    for m in data.get('missions',[])[:MAX_MISSIONS_TOTAL]:
        m.setdefault('entreprise',''); m.setdefault('poste','')
        m.setdefault('duree',''); m.setdefault('realisations',[])
        m['entreprise'] = fit_title_field(normalize_text(m.get('entreprise')), 55)
        m['poste'] = fit_title_field(normalize_text(m.get('poste')), 65)
        m['duree'] = normalize_text(m.get('duree'))[:40]
        if src_text:
            if m['entreprise'] and not grounded(m['entreprise'], src_text, src_tok, src_folded, min_ratio=0.5, min_hits=1):
                m['entreprise'] = ''
            if m['poste'] and not grounded(m['poste'], src_text, src_tok, src_folded, min_ratio=0.45, min_hits=1):
                m['poste'] = ''
            if m['duree'] and not grounded(m['duree'], src_text, src_tok, src_folded, min_ratio=1.0, min_hits=1):
                m['duree'] = ''
        reals = []
        for r in m['realisations'][:MAX_REALISATIONS_PER_MISSION]:
            r = fit_realisation(normalize_text(r), REALISATION_MAX_WORDS, REALISATION_MAX_CHARS)
            if not r:
                continue
            if src_text and not grounded(r, src_text, src_tok, src_folded, min_ratio=0.4, min_hits=2):
                continue
            reals.append(r)
        m['realisations'] = uniq(reals, MAX_REALISATIONS_PER_MISSION)
        if not m['entreprise'] and not m['poste'] and not m['realisations']:
            continue
        if not m['realisations']:
            fallback = normalize_text(' '.join(filter(None, [m['poste'], m['entreprise']]))).strip()
            m['realisations'] = [fit_realisation(fallback, REALISATION_MAX_WORDS, REALISATION_MAX_CHARS)] if fallback else []
        missions.append(m)

    derived_xp = derive_annees_xp(missions)
    if derived_xp and (not data['annees_xp'] or len(derived_xp) >= len(data['annees_xp'])):
        data['annees_xp'] = derived_xp

    if not titre:
        latest_poste = next((m.get('poste') for m in missions if m.get('poste')), '')
        titre = trunc(latest_poste, 6)[:50] if latest_poste else 'Consultant Finance'
    data['titre'] = titre or 'Consultant Finance'
    data['missions'] = missions
    return data

def merge_llm_outputs(primary, secondary, cv_text):
    """Conserve la version la plus fidèle et la plus riche entre les deux passes LLM."""
    if not isinstance(primary, dict):
        return secondary if isinstance(secondary, dict) else {}
    if not isinstance(secondary, dict):
        return primary

    src_text = normalize_text(cv_text)
    src_tok = source_tokens(src_text)
    src_folded = ascii_fold(src_text)

    merged = copy.deepcopy(primary)

    scalar_fields = ['prenom', 'nom', 'titre', 'annees_xp']
    list_fields = ['competences_cles', 'formation', 'outils', 'aptitudes']

    for field in scalar_fields:
        base = normalize_text(primary.get(field))
        alt = normalize_text(secondary.get(field))
        if alt and grounded(alt, src_text, src_tok, src_folded, min_ratio=0.4, min_hits=1):
            if not base or richer_text(alt, base):
                merged[field] = alt

    for field in list_fields:
        base_items = [normalize_text(x) for x in primary.get(field, []) if normalize_text(x)]
        alt_items = [normalize_text(x) for x in secondary.get(field, []) if normalize_text(x)]
        out = list(base_items)
        for item in alt_items:
            if not grounded(item, src_text, src_tok, src_folded, min_ratio=0.4, min_hits=1):
                continue
            replaced = False
            for i, existing in enumerate(out):
                if richer_text(item, existing):
                    out[i] = item
                    replaced = True
                    break
                if ascii_fold(item) == ascii_fold(existing):
                    replaced = True
                    break
            if not replaced:
                out.append(item)
        merged[field] = out

    base_missions = primary.get('missions', []) if isinstance(primary.get('missions'), list) else []
    alt_missions = secondary.get('missions', []) if isinstance(secondary.get('missions'), list) else []
    merged_missions = []
    for idx in range(max(len(base_missions), len(alt_missions))):
        base = copy.deepcopy(base_missions[idx]) if idx < len(base_missions) and isinstance(base_missions[idx], dict) else {}
        alt = alt_missions[idx] if idx < len(alt_missions) and isinstance(alt_missions[idx], dict) else {}
        current = base if base else copy.deepcopy(alt)
        for field in ['entreprise', 'poste', 'duree']:
            base_val = normalize_text(base.get(field))
            alt_val = normalize_text(alt.get(field))
            if alt_val and grounded(alt_val, src_text, src_tok, src_folded, min_ratio=0.4, min_hits=1):
                if not base_val or richer_text(alt_val, base_val):
                    current[field] = alt_val
                elif base_val:
                    current[field] = base_val
        base_reals = [normalize_text(x) for x in base.get('realisations', []) if normalize_text(x)]
        alt_reals = [normalize_text(x) for x in alt.get('realisations', []) if normalize_text(x)]
        reals = list(base_reals)
        for item in alt_reals:
            if not grounded(item, src_text, src_tok, src_folded, min_ratio=0.35, min_hits=2):
                continue
            replaced = False
            for i, existing in enumerate(reals):
                if richer_text(item, existing):
                    reals[i] = item
                    replaced = True
                    break
                if ascii_fold(item) == ascii_fold(existing):
                    replaced = True
                    break
            if not replaced:
                reals.append(item)
        current['realisations'] = reals
        merged_missions.append(current)
    merged['missions'] = merged_missions
    return merged

def process_cv(cv_text, debug=False):
    raw1 = call_llm([{'role':'system','content':PROMPT_EXTRACT},{'role':'user','content':f'CV:\n\n{cv_text}'}], temp=0.0, json_mode=True)
    parsed1 = parse_json(raw1, repair=True)
    data = parsed1
    raw2 = None
    d2 = None
    try:
        raw2 = call_llm([{'role':'user','content':PROMPT_VALID.format(cv_text=cv_text, json_data=json.dumps(data,ensure_ascii=False))}], temp=0.0, json_mode=True)
        d2 = parse_json(raw2, repair=True)
        if isinstance(d2.get('missions'),list) and d2.get('prenom'):
            data = merge_llm_outputs(data, d2, cv_text)
    except: pass
    data = merge_recovered_missions(data, cv_text)
    cleaned = clean(data, cv_text)
    if not debug:
        return cleaned
    return cleaned, {
        'pdf_text': cv_text,
        'raw_extract_text': raw1,
        'raw_extract_json': parsed1,
        'raw_valid_text': raw2,
        'raw_valid_json': d2,
        'final_json': cleaned,
    }

# ── ROUTES ────────────────────────────────────────────────────────────────────
@app.post("/api/generate")
async def generate(request: Request, file: UploadFile = File(...)):
    check_auth(request)
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, 'Fichier PDF requis')
    content = await file.read()
    if len(content) > 10*1024*1024:
        raise HTTPException(400, 'Fichier trop volumineux (max 10 Mo)')

    text = ''
    with pdfplumber.open(io.BytesIO(content)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: text += t + '\n\n'
    if not text.strip():
        raise HTTPException(400, "PDF non lisible — fournir un PDF avec texte sélectionnable")

    try: data, debug_payload = process_cv(text, debug=True)
    except Exception as e: raise HTTPException(500, f'Erreur LLM: {e}')

    ddc_id = str(uuid.uuid4())[:8]
    prenom = re.sub(r'[^a-zA-ZÀ-ÿ]','_', data.get('prenom','consultant'))
    filename = f"DDC_{prenom}_{ddc_id}.pptx"

    try: build_pptx(data, OUTPUT_DIR/filename)
    except Exception as e: raise HTTPException(500, f'Erreur PPT: {e}')

    entry = {
        'id':ddc_id,'filename':filename,
        'prenom':data.get('prenom',''),'nom':data.get('nom',''),
        'titre':data.get('titre',''),'annees_xp':data.get('annees_xp',''),
        'created_at':datetime.now().isoformat(),
        'cv_original':file.filename,'nb_missions':len(data.get('missions',[])),
        'competences_cles':data.get('competences_cles',[]),
        'outils':data.get('outils',[]),
        'missions':[{'entreprise':m.get('entreprise',''),'poste':m.get('poste',''),'duree':m.get('duree','')} for m in data.get('missions',[])],
    }
    history.insert(0, entry)
    _save_history(history)
    return {
        'ok':True,
        'id':ddc_id,
        'filename':filename,
        'data':data,
        'nb_missions':len(data.get('missions',[])),
        'debug':debug_payload,
    }

@app.get("/api/download/{filename}")
def download(filename: str, request: Request):
    check_auth(request)
    if '/' in filename or '..' in filename: raise HTTPException(400,'Nom invalide')
    path = OUTPUT_DIR/filename
    if not path.exists(): raise HTTPException(404,'Fichier non trouvé')
    return FileResponse(str(path), filename=filename,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation')

@app.get("/api/history")
def get_history(request: Request):
    check_auth(request)
    return history

@app.get("/api/status")
def status():
    try:
        client = ollama.Client(host=OLLAMA_HOST)
        models = client.list()
        names = [m['name'] for m in models.get('models',[])]
        return {'ollama':True,'model':MODEL,'model_ready':any(MODEL in m for m in names)}
    except: return {'ollama':False,'model':MODEL,'model_ready':False}
