import os, sys, json, io, tempfile, datetime
import anthropic
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
sys.path.insert(0, '/root')
from analyze_ao import analyze_ao_gonogo
from gonogo_report import generate_gonogo_report
from planning_agent import run_planning_agent

# Args from n8n
FILE_ID = sys.argv[1] if len(sys.argv) > 1 else None
FILE_NAME = sys.argv[2] if len(sys.argv) > 2 else "AO_inconnu"

print(f"🚀 Pipeline v2 started for: {FILE_NAME} ({FILE_ID})")

SERVICE_ACCOUNT_FILE = '/root/chg-credentials/service-account.json'
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=['https://www.googleapis.com/auth/drive'])
drive = build('drive', 'v3', credentials=creds)
claude = anthropic.Anthropic()

cvs_raw = json.load(open('/root/chg-library/cv_docx_library.json'))
cvs = [{'nom': cv.get('nom', ''), 'titre': cv.get('nom', ''), 'competences': [], 'text': cv.get('texte', '')} for cv in cvs_raw]
projects = json.load(open('/root/chg-library/projects_library.json'))
real_proposal = json.load(open('/root/chg-library/real_proposal_analysis.json'))
crossref = json.load(open('/root/chg-library/ao_os_crossref.json'))
crossref_intel = {'lecons': [], 'arguments_differenciateurs': [], 'mots_cles_techniques': [], 'mappings': []}
for pair in crossref.get('pairs', []):
    crossref_intel['lecons'].extend(pair.get('lecons_pour_futurs_aos', []))
    crossref_intel['arguments_differenciateurs'].extend(pair.get('arguments_differenciateurs', []))
    crossref_intel['mots_cles_techniques'].extend(pair.get('mots_cles_techniques', []))
    crossref_intel['mappings'].extend(pair.get('comment_chg_repond_aux_exigences', []))
crossref_intel['lecons'] = list(dict.fromkeys(crossref_intel['lecons']))[:20]
crossref_intel['arguments_differenciateurs'] = list(dict.fromkeys(crossref_intel['arguments_differenciateurs']))[:10]
crossref_intel['mots_cles_techniques'] = list(dict.fromkeys(crossref_intel['mots_cles_techniques']))[:15]

def download_bytes(file_id):
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, drive.files().get_media(fileId=file_id))
    done = False
    while not done: _, done = dl.next_chunk()
    return buf.getvalue()

def pdf_to_text(pdf_bytes, max_chars=10000):
    try:
        from pdfminer.high_level import extract_text as extract_text_pdf
        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
            f.write(pdf_bytes); tmp = f.name
        text = extract_text_pdf(tmp); os.unlink(tmp)
        return text[:max_chars]
    except: return ""

def docx_to_text(docx_bytes, max_chars=10000):
    import io as _io
    from docx import Document as _DocxDoc
    try:
        doc = _DocxDoc(_io.BytesIO(docx_bytes))
        text = chr(10).join([p.text for p in doc.paragraphs if p.text.strip()])
        return text[:max_chars]
    except Exception as e:
        print(f"Warning: docx extraction error: {e}")
        return ""

def claude_call(prompt, max_tokens=2000):
    r = claude.messages.create(model="claude-sonnet-4-5-20250929", max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}])
    resp = r.content[0].text.strip()
    if resp.startswith("```"): resp = resp.split("\n",1)[1].rsplit("```",1)[0].strip()
    return resp

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        run.font.bold = True

def add_divider(doc):
    p = doc.add_paragraph('—' * 80)
    p.runs[0].font.size = Pt(6)
    p.runs[0].font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

def match_cv(member_name, cvs):
    member_lower = member_name.lower()
    member_base = member_lower.split(' - ')[0].strip()
    member_words = set(member_base.split())
    best_cv = None
    best_score = 0
    for c in cvs:
        cv_base = c['nom'].lower().split(' - ')[0].strip()
        cv_words = set(cv_base.split())
        score = len(member_words & cv_words)
        if score > best_score:
            best_score = score
            best_cv = c
    return best_cv if best_score >= 1 else None

# ── STEP 1: Download and extract AO text ──────────────────────────────────────
print("📄 Downloading AO...")
ao_bytes = download_bytes(FILE_ID)

def docx_to_text_inner(docx_bytes, max_chars=10000):
    import io as _io
    from docx import Document as _D
    try:
        doc = _D(_io.BytesIO(docx_bytes))
        return chr(10).join([p.text for p in doc.paragraphs if p.text.strip()])[:max_chars]
    except: return ""

is_docx = FILE_NAME.lower().endswith('.docx') or FILE_NAME.lower().endswith('.doc')
if is_docx:
    ao_pdf_tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
    ao_pdf_tmp.write(ao_bytes); ao_pdf_tmp.close()
    ao_pdf_path = ao_pdf_tmp.name
    ao_text = docx_to_text_inner(ao_bytes)
    print(f"📝 Detected Word document — extracting text from .docx")
else:
    ao_pdf_tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    ao_pdf_tmp.write(ao_bytes); ao_pdf_tmp.close()
    ao_pdf_path = ao_pdf_tmp.name
    ao_text = pdf_to_text(ao_bytes)[:10000]
    print(f"📄 Detected PDF — extracting text with pdfminer")

# ── STEP 2: Analyze AO ────────────────────────────────────────────────────────
print("🧠 Analyzing AO with Claude...")
r = claude.messages.create(model="claude-sonnet-4-5-20250929", max_tokens=2000,
    messages=[{"role": "user", "content": f"AO:\n{ao_text}\n\nUNIQUEMENT JSON sans backticks:\n{{\"client\":\"\",\"numero_ao\":\"\",\"titre_projet\":\"\",\"date_limite\":\"\",\"type_travaux\":\"\",\"competences_requises\":[],\"localisation\":\"\",\"resume_mandat\":\"\",\"points_addendas\":[]}}"}])
resp = r.content[0].text.strip()
if resp.startswith("```"): resp = resp.split("\n",1)[1].rsplit("```",1)[0].strip()
ao_data = json.loads(resp)
print(f"DEBUG ao_data keys: {list(ao_data.keys())}")
print(f"✅ AO: {ao_data['client']} — {ao_data['titre_projet']}")

# ── STEP 3: Go/No-Go Analysis ─────────────────────────────────────────────────
print("\n🔍 Running Go/No-Go analysis...")
gonogo_score = 0; gonogo_recommandation = ""; gonogo_nb_incongruites = 0; gonogo_report_path = None; planning_report_path = None
try:
    gonogo_data = analyze_ao_gonogo(ao_pdf_path)
    gonogo_score = gonogo_data.get('score_gonogo', 0)
    gonogo_recommandation = gonogo_data.get('recommandation', '')
    gonogo_nb_incongruites = len(gonogo_data.get('incongruites', []))
    print(f"✅ Go/No-Go: {gonogo_score}/100 — {gonogo_recommandation}")
    gonogo_report_path = f"/local-files/LATEST_GONOGO_RAPPORT.docx"
    import tempfile as _tf
    _tmp_json = _tf.NamedTemporaryFile(suffix='.json', delete=False, mode='w')
    json.dump(gonogo_data, _tmp_json, ensure_ascii=False)
    _tmp_json.close()
    generate_gonogo_report(_tmp_json.name, gonogo_report_path)
    os.unlink(_tmp_json.name)
    print(f"✅ Go/No-Go report: {gonogo_report_path}")


except Exception as e:
    print(f"⚠️ Go/No-Go failed (non-blocking): {e}")


# ── Derive safe_name from FILE_NAME (same logic as pipeline.py) ──────────────
import re as _re
safe_name = FILE_NAME
for ext in ['.pdf', '.docx', '.doc']:
    safe_name = safe_name.replace(ext, '').replace(ext.upper(), '')
safe_name = safe_name.replace('Proposition_CHG_', '')
safe_name = _re.sub(r'[^a-zA-Z0-9_\-]', '_', safe_name)[:60]

# ── Agent 2 — Planning document ──────────────────────────────────────────────
planning_report_path = None
try:
    import tempfile as _tf2, json as _json2
    planning_report_path = f"/local-files/PLANIFICATION_{safe_name}.docx"
    with _tf2.NamedTemporaryFile(suffix=".json", delete=False, mode="w", encoding="utf-8") as _tmp2:
        _json2.dump(gonogo_data, _tmp2, ensure_ascii=False)
        _tmp2_name = _tmp2.name
    run_planning_agent(_tmp2_name, planning_report_path)
    os.unlink(_tmp2_name)
    print(f"\u2705 Planning document: {planning_report_path}")
except Exception as _ep:
    print(f"\u26a0\ufe0f Planning agent error: {_ep}")
    planning_report_path = None

# ── N8N_OUTPUT ────────────────────────────────────────────────────────────────
result = {
    "status": "success",
    "ao_client": ao_data.get("client", ao_data.get("municipality", ao_data.get("municipalite", ""))),
    "ao_titre": ao_data.get("titre_projet", ao_data.get("titre", ao_data.get("title", ""))),
    "gonogo_score": gonogo_score,
    "gonogo_recommandation": gonogo_recommandation,
    "gonogo_nb_incongruites": gonogo_nb_incongruites,
    "gonogo_report_file": gonogo_report_path.replace("/local-files/", "/files/") if gonogo_report_path else "",
    "planning_report_file": planning_report_path.replace("/local-files/", "/files/") if planning_report_path else "",
}
print(f"N8N_OUTPUT:{json.dumps(result)}")
