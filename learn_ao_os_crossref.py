import os
import json
import zipfile
import re
import anthropic
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
from io import BytesIO, StringIO
from docx import Document
from pdfminer.high_level import extract_text_to_fp
from pdfminer.layout import LAParams

# --- Config ---
OS_EXEMPLES_FOLDER_ID = '11rz8H6p378Q6mgcQYVeEAM9PNW7KSGEg'
OUTPUT_FILE = '/root/chg-library/ao_os_crossref.json'
MAX_CHARS = 50000

# --- Clients ---
creds = service_account.Credentials.from_service_account_file(
    '/root/chg-credentials/service-account.json',
    scopes=['https://www.googleapis.com/auth/drive.readonly']
)
drive = build('drive', 'v3', credentials=creds)
claude = anthropic.Anthropic(api_key=os.environ['ANTHROPIC_API_KEY'])

def download_bytes(file_id):
    request = drive.files().get_media(fileId=file_id)
    buf = BytesIO()
    dl = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()

def extract_word(file_bytes):
    try:
        doc = Document(BytesIO(file_bytes))
        sections = []
        for element in doc.element.body:
            tag = element.tag.split('}')[-1]
            if tag == 'p':
                text = ''.join([n.text or '' for n in element.iter() if n.tag.endswith('}t')])
                if text.strip():
                    sections.append(text.strip())
            elif tag == 'tbl':
                rows = []
                for row in element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'):
                    cells = []
                    for cell in row.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc'):
                        ct = ''.join([n.text or '' for n in cell.iter() if n.tag.endswith('}t')]).strip()
                        if ct:
                            cells.append(ct)
                    if cells:
                        rows.append(' | '.join(cells))
                if rows:
                    sections.append('\n'.join(rows))
        return '\n'.join(sections)
    except Exception:
        # docm fallback — XML extraction
        with zipfile.ZipFile(BytesIO(file_bytes)) as z:
            if 'word/document.xml' in z.namelist():
                xml = z.read('word/document.xml').decode('utf-8', errors='ignore')
                return ' '.join(re.findall(r'<w:t[^>]*>([^<]+)</w:t>', xml))
        return ''

def extract_pdf(file_bytes):
    output = StringIO()
    try:
        extract_text_to_fp(BytesIO(file_bytes), output, laparams=LAParams())
    except Exception:
        pass
    return output.getvalue().strip()

def truncate(text, max_chars=MAX_CHARS):
    if len(text) <= max_chars:
        return text
    keep_start = int(max_chars * 0.6)
    keep_end = int(max_chars * 0.4)
    return text[:keep_start] + '\n\n[...]\n\n' + text[-keep_end:]

def get_folder_pairs():
    """Get AO + OS file pairs from each subfolder"""
    subfolders = drive.files().list(
        q=f"'{OS_EXEMPLES_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)"
    ).execute().get('files', [])

    pairs = []
    for folder in subfolders:
        files = drive.files().list(
            q=f"'{folder['id']}' in parents",
            fields="files(id, name, mimeType, size)"
        ).execute().get('files', [])

        ao_file = None
        os_file = None

        for f in files:
            name_lower = f['name'].lower()
            is_word = 'wordprocessing' in f['mimeType'] or f['name'].endswith('.docm') or f['name'].endswith('.docx')
            is_pdf = f['mimeType'] == 'application/pdf'

            # Skip bordereaux and financial offers
            if any(skip in name_lower for skip in ['bordereau', 'financ']):
                continue

            if is_word:
                os_file = f  # Word file = CHG's OS response
            elif is_pdf and ao_file is None:
                # Largest PDF = main AO document
                if ao_file is None or int(f.get('size', 0)) > int(ao_file.get('size', 0)):
                    ao_file = f

        if ao_file and os_file:
            pairs.append({
                'folder': folder['name'],
                'ao': ao_file,
                'os': os_file
            })
            print(f"   ✅ Pair found: {folder['name']}")
            print(f"      AO: {ao_file['name']}")
            print(f"      OS: {os_file['name']}")
        else:
            print(f"   ⚠️  Incomplete pair in {folder['name']} — AO: {ao_file is not None}, OS: {os_file is not None}")

    return pairs

def analyze_pair(pair):
    """Cross-analyze AO + OS together to extract strategic mapping"""
    print(f"\n🔍 Cross-analyzing: {pair['folder']}...")

    # Download and extract both documents
    print(f"   Downloading AO...")
    ao_bytes = download_bytes(pair['ao']['id'])
    ao_text = truncate(extract_pdf(ao_bytes))
    print(f"   AO text: {len(ao_text)} chars")

    print(f"   Downloading OS...")
    os_bytes = download_bytes(pair['os']['id'])
    os_text = truncate(extract_word(os_bytes))
    print(f"   OS text: {len(os_text)} chars")

    prompt = f"""Tu es un expert en analyse de propositions d'ingénierie québécoises.

Voici un appel d'offres (AO) et la réponse gagnante du Groupe Conseil CHG (OS).
Analyse les deux documents ensemble pour extraire la stratégie de réponse de CHG.

=== APPEL D'OFFRES (AO) ===
{ao_text}

=== OFFRE DE SERVICE CHG (OS) ===
{os_text}

Retourne UNIQUEMENT un JSON valide sans backticks:
{{
  "projet": "",
  "type_mandat": "",
  "exigences_cles_ao": [],
  "comment_chg_repond_aux_exigences": [
    {{"exigence": "", "reponse_chg": "", "formule_utilisee": ""}}
  ],
  "arguments_differenciateurs": [],
  "structure_reponse_chg": [],
  "ton_adapte_au_client": "",
  "phrases_ouverture_gagnantes": [],
  "phrases_cloture_gagnantes": [],
  "mots_cles_techniques": [],
  "points_forts_mis_en_avant": [],
  "lecons_pour_futurs_aos": []
}}"""

    response = claude.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )

    text = response.content[0].text.strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    start = text.find('{')
    end = text.rfind('}') + 1
    if start >= 0 and end > start:
        text = text[start:end]

    try:
        result = json.loads(text)
    except json.JSONDecodeError:
        print(f"   ⚠️  JSON error, retrying...")
        response2 = claude.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=2000,
            messages=[{"role": "user", "content": f"Retourne UNIQUEMENT ce JSON valide pour le projet {pair['folder']}:\n{{\"projet\": \"\", \"type_mandat\": \"\", \"exigences_cles_ao\": [], \"comment_chg_repond_aux_exigences\": [], \"arguments_differenciateurs\": [], \"lecons_pour_futurs_aos\": []}}"}]
        )
        t2 = response2.content[0].text.strip()
        s = t2.find('{'); e = t2.rfind('}') + 1
        result = json.loads(t2[s:e])

    result['folder'] = pair['folder']
    result['ao_source'] = pair['ao']['name']
    result['os_source'] = pair['os']['name']
    return result

# --- Main ---
print("🚀 AO/OS Cross-Reference Learning...\n")
pairs = get_folder_pairs()
print(f"\n📊 Found {len(pairs)} AO/OS pairs to analyze\n")

analyses = []
for pair in pairs:
    try:
        analysis = analyze_pair(pair)
        analyses.append(analysis)
        print(f"   ✅ Done: {pair['folder']}")
    except Exception as e:
        print(f"   ❌ Failed: {pair['folder']} — {e}")

# Save output
output = {
    "total_pairs_analyzed": len(analyses),
    "pairs": analyses
}

with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"\n✅ Cross-reference analysis saved to {OUTPUT_FILE}")
print(f"📚 Pairs analyzed: {len(analyses)}")
for a in analyses:
    print(f"   - {a['folder']}")
