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

# ── STEP 4: Select team ───────────────────────────────────────────────────────
cv_summary = [{"nom": c["nom"], "titre": c.get("titre",""), "competences": c.get("competences",[])} for c in cvs]
r = claude.messages.create(model="claude-sonnet-4-5-20250929", max_tokens=800,
    messages=[{"role": "user", "content": f"AO: {json.dumps(ao_data, ensure_ascii=False)}\nRessources: {json.dumps(cv_summary, ensure_ascii=False)}\nSélectionne 5-6 ressources. UNIQUEMENT JSON sans backticks:\n{{\"equipe\":[{{\"nom\":\"\",\"role_dans_projet\":\"\"}}]}}"}])
resp = r.content[0].text.strip()
if resp.startswith("```"): resp = resp.split("\n",1)[1].rsplit("```",1)[0].strip()
team_data = json.loads(resp)
print(f"✅ Team: {[m['nom'] for m in team_data['equipe']]}")

# ── STEP 5: Select and parse reference projects ───────────────────────────────
print("📂 Parsing reference projects...")

def select_relevant_projects(projects, ao_data, n=5):
    """Select most relevant projects based on AO type keywords."""
    import re
    ao_text_lower = (ao_data.get('type_travaux','') + ' ' + ao_data.get('titre_projet','')).lower()
    keywords = re.findall(r'\b\w{4,}\b', ao_text_lower)
    scored = []
    for p in projects:
        texte = p.get('texte','').lower()
        nom = p.get('nom','').lower()
        score = sum(1 for kw in keywords if kw in texte or kw in nom)
        scored.append((score, p))
    scored.sort(key=lambda x: -x[0])
    return [p for _, p in scored[:n]]

def parse_project_regex(proj):
    """Extract structured fields from fiche text using regex/heuristics — no API call."""
    import re
    texte = proj.get('texte', '')
    nom_fiche = proj.get('nom', proj.get('fichier', 'Projet'))
    
    # Extract project name from first meaningful line
    lines = [l.strip() for l in texte.split('\n') if l.strip() and len(l.strip()) > 5]
    nom_projet = lines[0][:80] if lines else nom_fiche
    
    # Extract client (line after "Client" label or before pipe separator)
    client = ''
    for line in lines:
        if 'client' in line.lower() and '|' in line:
            parts = line.split('|')
            for i, p in enumerate(parts):
                if 'client' in p.lower() and i+1 < len(parts):
                    client = parts[i+1].strip()[:60]
                    break
        if not client:
            m = re.search(r'(?:Municipalit[eé]|Ville|MRC|Office|Soci[eé]t[eé]|Commission|[A-Z][a-z]+ de [A-Z])[^|\n]{3,40}', texte)
            if m: client = m.group(0).strip()[:60]
    
    # Extract year from fiche number (e.g. "22-171" → 2022)
    m = re.match(r'(\d{2})-', nom_fiche)
    annee = f"20{m.group(1)}" if m else ''
    
    # Extract honoraires
    valeur = ''
    m = re.search(r'(\d+[,.]?\d*\s*[kKmM]?\$)', texte)
    if m: valeur = m.group(1)
    
    # Extract services from description
    desc_match = re.search(r'Description du projet[^|\n]*[|:]([^|\n]{20,200})', texte, re.IGNORECASE)
    services = desc_match.group(1).strip()[:200] if desc_match else texte[:200]
    
    return {
        'nom_projet': nom_projet,
        'client': client,
        'localisation': '',
        'annee': annee,
        'valeur_contrat': valeur,
        'services_rendus': services,
        'description_courte': texte[:300]
    }

selected_projects_raw = select_relevant_projects(projects, ao_data, n=5)
parsed_projects = []
for proj in selected_projects_raw:
    parsed = parse_project_regex(proj)
    parsed_projects.append(parsed)
    print(f"  ✅ {parsed.get('nom_projet','?')[:50]} | client: {parsed.get('client','?')[:30]} | {parsed.get('annee','')} | {parsed.get('valeur_contrat','')}")

# ── STEP 6: Detect AO evaluation grid ────────────────────────────────────────
print("📋 Detecting AO evaluation grid...")
try:
    grid_prompt = f"""Analyse cet appel d'offres et identifie les sections d'évaluation exactes utilisées par le comité de sélection.

AO:
{ao_text[:4000]}

UNIQUEMENT JSON sans backticks:
{{
  "sections_evaluation": [
    {{"numero": "1", "titre": "Expérience du soumissionnaire", "poids_pct": 30, "criteres": ["sous-critère 1", "sous-critère 2"]}},
    {{"numero": "2", "titre": "Expérience du chargé de projet", "poids_pct": 25, "criteres": []}},
    {{"numero": "3", "titre": "Organisation et expérience de l'équipe", "poids_pct": 25, "criteres": []}},
    {{"numero": "4", "titre": "Compréhension du mandat et méthodologie", "poids_pct": 20, "criteres": []}}
  ],
  "structure_detectee": "standard_4_sections"
}}

Si l'AO ne précise pas de grille, utilise la structure standard CHG à 4 sections ci-dessus."""
    grid_resp = claude_call(grid_prompt, max_tokens=1000)
    eval_grid = json.loads(grid_resp)
    print(f"✅ Eval grid: {len(eval_grid['sections_evaluation'])} sections")
except Exception as e:
    print(f"⚠️ Grid detection failed, using default: {e}")
    eval_grid = {"sections_evaluation": [
        {"numero": "1", "titre": "Expérience du soumissionnaire", "poids_pct": 30, "criteres": ["Expérience en travaux similaires", "Contrats de référence"]},
        {"numero": "2", "titre": "Expérience du chargé de projet", "poids_pct": 25, "criteres": ["Formation et titre professionnel", "Projets similaires réalisés", "Disponibilité"]},
        {"numero": "3", "titre": "Organisation et expérience de l'équipe", "poids_pct": 25, "criteres": ["Composition de l'équipe", "Expérience des membres", "Capacité de relève"]},
        {"numero": "4", "titre": "Compréhension du mandat et méthodologie", "poids_pct": 20, "criteres": ["Compréhension des enjeux", "Méthodologie détaillée", "Échéancier"]}
    ], "structure_detectee": "standard_4_sections"}

# ── STEP 7: Generate all content sections ────────────────────────────────────
print("✍️  Generating proposal content with Claude...")

# Find the chargé de projet (first team member or most senior)
charge_projet = team_data['equipe'][0] if team_data['equipe'] else {"nom": "Charles Gauthier", "role_dans_projet": "Chargé de projet"}
charge_cv = match_cv(charge_projet['nom'], cvs)
charge_cv_text = charge_cv.get('cv_text', charge_cv.get('text', ''))[:2000] if charge_cv else ""

autres_membres = team_data['equipe'][1:] if len(team_data['equipe']) > 1 else []
autres_cvs_summary = []
for m in autres_membres:
    cv = match_cv(m['nom'], cvs)
    if cv:
        autres_cvs_summary.append(f"{m['nom']} ({m['role_dans_projet']}): {cv.get('cv_text', cv.get('text',''))[:500]}")

projects_summary = json.dumps([{
    'nom': p.get('nom_projet',''), 'client': p.get('client',''),
    'annee': p.get('annee',''), 'valeur': p.get('valeur_contrat',''),
    'services': p.get('services_rendus','')[:200]
} for p in parsed_projects], ensure_ascii=False)

sections_list = "\n".join([f"  Section {s['numero']} ({s['poids_pct']}%): {s['titre']}" for s in eval_grid['sections_evaluation']])

content_prompt = f"""Tu es un expert senior en génie-conseil québécois qui rédige des offres de service gagnantes pour le Groupe Conseil CHG inc.

TON OBJECTIF : Produire un OS qui GAGNE — pas juste qui répond à l'AO. Chaque section doit maximiser le score selon la grille d'évaluation.

=== APPEL D'OFFRES ===
{json.dumps(ao_data, ensure_ascii=False)}

=== GRILLE D'ÉVALUATION (sections scorées par l'évaluateur) ===
{sections_list}

=== CHARGÉ DE PROJET ===
{charge_projet['nom']} — {charge_projet['role_dans_projet']}
CV: {charge_cv_text}

=== AUTRES MEMBRES DE L'ÉQUIPE ===
{chr(10).join(autres_cvs_summary[:4])}

=== PROJETS DE RÉFÉRENCE SÉLECTIONNÉS ===
{projects_summary}

=== STYLE D'ÉCRITURE CHG (reproduire exactement) ===
Ton et style: {real_proposal.get('ton_et_style', '')}
Formules d'ouverture: {json.dumps(real_proposal.get('formules_ouverture', [])[:4], ensure_ascii=False)}
Phrases clés CHG: {json.dumps(real_proposal.get('phrases_cles', [])[:8], ensure_ascii=False)}

=== INTELLIGENCE STRATÉGIQUE (leçons de vraies propositions gagnantes) ===
Leçons: {json.dumps(crossref_intel['lecons'][:8], ensure_ascii=False)}
Arguments différenciateurs: {json.dumps(crossref_intel['arguments_differenciateurs'][:6], ensure_ascii=False)}
Mots clés techniques: {json.dumps(crossref_intel['mots_cles_techniques'][:10], ensure_ascii=False)}

=== INSTRUCTIONS DE RÉDACTION ===
1. Écris comme un ingénieur CHG expérimenté — professionnel, précis, confiant sans arrogance
2. Chaque section doit répondre DIRECTEMENT aux critères de la grille d'évaluation
3. Section 1 (Expérience firme): mets en avant les contrats similaires, la taille de la firme, certifications ISO 9001:2015
4. Section 2 (Chargé de projet): présentation complète avec formation, OIQ, projets pertinents, disponibilité explicite
5. Section 3 (Équipe): présente chaque membre avec son rôle spécifique + expérience pertinente + capacité de relève
6. Section 4 (Méthodologie): phases détaillées adaptées à CE projet spécifique, pas générique
7. Intègre les arguments différenciateurs naturellement — ne pas les lister
8. NE PAS utiliser de langage générique ou formules creuses
9. Chaque section : minimum 300 mots, maximum 500 mots

RETOURNE UNIQUEMENT un JSON valide sans backticks:
{{
  "lettre_presentation": "lettre formelle 4 paragraphes, ton CHG, adressée au client spécifique",
  "section_1_experience_firme": "texte complet section 1 — expérience du soumissionnaire",
  "section_2_charge_projet": "texte complet section 2 — chargé de projet avec détails CV",
  "section_3_equipe": "texte complet section 3 — organisation et expérience équipe",
  "section_4_methodologie": "texte complet section 4 — compréhension mandat et méthodologie détaillée par phases",
  "conclusion": "mot de la fin engageant 100 mots style CHG"
}}"""

r = claude.messages.create(model="claude-sonnet-4-5-20250929", max_tokens=8000,
    messages=[{"role": "user", "content": content_prompt}])
resp = r.content[0].text.strip()
if resp.startswith("```"): resp = resp.split("\n",1)[1].rsplit("```",1)[0].strip()
sections = json.loads(resp)
print("✅ Sections generated")

# ── STEP 8: Build Word document ───────────────────────────────────────────────
print("📝 Building Word document...")
doc = Document()
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# ── Cover page ────────────────────────────────────────────────────────────────
doc.add_paragraph("")
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("GROUPE CONSEIL CHG INC.")
run.bold = True; run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
doc.add_paragraph("")
p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run2 = p2.add_run("PROPOSITION DE SERVICES PROFESSIONNELS")
run2.bold = True; run2.font.size = Pt(16)
doc.add_paragraph("")
add_divider(doc)
doc.add_paragraph("")
for label, value in [
    ("PRÉSENTÉE À", ao_data['client'].upper()),
    ("PROJET", ao_data['titre_projet']),
    ("NUMÉRO AO", ao_data.get('numero_ao','')),
    ("DATE", datetime.date.today().strftime("%d %B %Y").lstrip("0")),
    ("NOTRE RÉFÉRENCE", f"CHG-{ao_data.get('numero_ao','').replace('-','')[:8]}"),
]:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"{label} : ").bold = True
    p.add_run(value)
doc.add_paragraph("")
add_divider(doc)
doc.add_page_break()

# ── Confidentiality clause ────────────────────────────────────────────────────
add_heading(doc, "AVIS DE CONFIDENTIALITÉ", 1)
p = doc.add_paragraph(
    "Cette offre de service est la propriété du Groupe Conseil CHG. Il y est fait état du savoir-faire "
    "de l'entreprise et des éléments qui la distinguent de ses compétiteurs. En conséquence, les "
    "informations contenues dans ce document sont strictement confidentielles et ne doivent être "
    "communiquées qu'aux personnes directement impliquées dans le processus d'évaluation. Toute "
    "reproduction ou diffusion, en tout ou en partie, est interdite sans le consentement écrit du "
    "Groupe Conseil CHG."
)
p.italic = True
doc.add_page_break()

# ── Table des matières ───────────────────────────────────────────────────────
add_heading(doc, "TABLE DES MATIÈRES", 1)
toc_entries = [
    ("Lettre de présentation", ""),
]
# Build TOC from eval grid sections
for es in eval_grid['sections_evaluation']:
    toc_entries.append((f"{es['numero']}   {es['titre']}", f"{es['poids_pct']}%"))
toc_entries += [
    ("Annexe 1 — Curriculum vitae", ""),
    ("Annexe 2 — Documents administratifs", ""),
]
for title, note in toc_entries:
    p = doc.add_paragraph()
    p.add_run(title)
    if note:
        p.add_run(f"  ({note})").italic = True
doc.add_paragraph("")
p = doc.add_paragraph()
p.add_run("Note : Mettre à jour la table des matières dans Word via Références → Mettre à jour la table.").italic = True
p.runs[0].font.size = Pt(9)
doc.add_page_break()

# ── Lettre de présentation ────────────────────────────────────────────────────
add_heading(doc, "LETTRE DE PRÉSENTATION", 1)
for para in sections['lettre_presentation'].split('\n\n'):
    if para.strip(): doc.add_paragraph(para.strip())
doc.add_page_break()

# ── Scored sections (matching AO eval grid) ───────────────────────────────────
section_map = {
    "1": ("section_1_experience_firme", parsed_projects),
    "2": ("section_2_charge_projet", None),
    "3": ("section_3_equipe", None),
    "4": ("section_4_methodologie", None),
}

for eval_sec in eval_grid['sections_evaluation']:
    num = eval_sec['numero']
    titre = eval_sec['titre'].upper()
    poids = eval_sec['poids_pct']
    criteres = eval_sec.get('criteres', [])

    heading_text = f"{num}.   {titre}  ({poids}%)"
    add_heading(doc, heading_text, 1)

    # Add criteria note (what evaluator scores) — subtle, helps reader see alignment
    if criteres:
        p = doc.add_paragraph()
        run = p.add_run(f"Critères d'évaluation : {' · '.join(criteres)}")
        run.italic = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        doc.add_paragraph("")

    sec_index = eval_grid['sections_evaluation'].index(eval_sec)
    position_map = {0: "section_1_experience_firme", 1: "section_2_charge_projet", 2: "section_3_equipe", 3: "section_4_methodologie"}
    content_key = position_map.get(sec_index)
    if content_key and content_key in sections:
        for para in sections[content_key].split('\n\n'):
            if para.strip(): doc.add_paragraph(para.strip())

    # Section 1: add reference project tables
    if sec_index == 0 and parsed_projects:
        doc.add_paragraph("")
        add_heading(doc, "Contrats de référence similaires", 2)
        for i, proj in enumerate(parsed_projects, 1):
            add_heading(doc, f"{num}.{i}   {proj.get('nom_projet', 'Projet')[:80]}", 3)
            table = doc.add_table(rows=5, cols=2)
            table.style = 'Table Grid'
            fields = [
                ("Client", proj.get('client', '')),
                ("Localisation", proj.get('localisation', '')),
                ("Année", proj.get('annee', '')),
                ("Valeur du contrat", proj.get('valeur_contrat', '')),
                ("Services rendus", proj.get('services_rendus', '')[:300]),
            ]
            for row_idx, (label, value) in enumerate(fields):
                row = table.rows[row_idx]
                row.cells[0].text = label
                row.cells[0].paragraphs[0].runs[0].bold = True
                row.cells[1].text = value
            doc.add_paragraph("")

    # Section 2: add chargé de projet availability box
    if sec_index == 1:
        doc.add_paragraph("")
        p = doc.add_paragraph()
        p.add_run("Disponibilité : ").bold = True
        p.add_run(f"{charge_projet['nom']} est disponible pour ce mandat dès la signature du contrat et s'engage à y consacrer le temps nécessaire à la bonne réalisation des travaux.")

    # Section 3: add team summary table
    if sec_index == 2:
        doc.add_paragraph("")
        add_heading(doc, "Composition de l'équipe", 2)
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        headers = ["Nom", "Titre", "Rôle dans le projet"]
        for i, h in enumerate(headers):
            table.rows[0].cells[i].text = h
            table.rows[0].cells[i].paragraphs[0].runs[0].bold = True
        for member in team_data['equipe']:
            cv = match_cv(member['nom'], cvs)
            row = table.add_row().cells
            row[0].text = member['nom']
            row[1].text = cv.get('titre', cv.get('nom', '')) if cv else ''
            row[2].text = member['role_dans_projet']
        doc.add_paragraph("")

    doc.add_page_break()

# ── Annexe A — CVs ────────────────────────────────────────────────────────────
add_heading(doc, "ANNEXE 1 — CURRICULUM VITAE", 1)
for member in team_data['equipe']:
    cv = match_cv(member['nom'], cvs)
    if cv:
        add_heading(doc, cv['nom'], 2)
        for line in cv.get('cv_text', cv.get('text', '')).split('\n'):
            if line.strip():
                doc.add_paragraph(line.strip())
        doc.add_page_break()

# ── Annexe B — Admin docs ─────────────────────────────────────────────────────
add_heading(doc, "ANNEXE 2 — DOCUMENTS ADMINISTRATIFS", 1)
for doc_name, desc in [
    ("Attestation de Revenu Québec", "Conformité fiscale"),
    ("Assurance responsabilité civile", "2 000 000 $ par sinistre"),
    ("Assurance responsabilité professionnelle", "2 000 000 $ par sinistre"),
    ("Attestation CNESST", "Conformité SST"),
    ("Certification ISO 9001:2015", "Valide jusqu'en 2026"),
    ("Déclaration d'absence de collusion", "Conformément aux exigences"),
]:
    p = doc.add_paragraph()
    p.add_run(f"• {doc_name} : ").bold = True
    p.add_run(desc)

# ── Mot de la fin ─────────────────────────────────────────────────────────────
add_heading(doc, "MOT DE LA FIN", 1)
for para in sections['conclusion'].split('\n\n'):
    if para.strip(): doc.add_paragraph(para.strip())

# ── Save ──────────────────────────────────────────────────────────────────────
safe_name = FILE_NAME.replace('.pdf','').replace('.docx','').replace(' ','_').replace('Proposition_CHG_','').replace('(','').replace(')','').replace('[','').replace(']','').replace(',','').replace(';','')[:40]
output_path = f"/local-files/Proposition_CHG_{safe_name}.docx"
doc.save(output_path)

# ── Agent 2 — Planning document ─────────────────────────────────────────
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



print(f"✅ Document saved: {output_path}")

drive_link = ''

result = {
    "status": "success",
    "ao_client": ao_data.get('client', ao_data.get('municipality', ao_data.get('municipalite', ''))),
    "ao_titre": ao_data.get('titre_projet', ao_data.get('titre', ao_data.get('title', ''))),
    "output_file": output_path.replace("/local-files/", "/files/"),
    "drive_link": drive_link,
    "team": [m['nom'] for m in team_data['equipe']],
    "gonogo_score": gonogo_score,
    "gonogo_recommandation": gonogo_recommandation,
    "gonogo_nb_incongruites": gonogo_nb_incongruites,
    "gonogo_report_file": gonogo_report_path.replace("/local-files/", "/files/") if gonogo_report_path else "",
        "planning_report_file": planning_report_path.replace("/local-files/", "/files/") if planning_report_path else ""
}
try: os.unlink(ao_pdf_path)
except: pass
print(f"N8N_OUTPUT:{json.dumps(result)}")

# ── LOCAL TEST MODE ───────────────────────────────────────────────────────────
# Usage: python3 pipeline_v2.py LOCAL /local-files/AO_file.pdf "AO_name.pdf"
# Overrides download_bytes to read from local path instead of Drive
