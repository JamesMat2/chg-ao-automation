import anthropic
import json
import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from lxml import etree
from docx.oxml import OxmlElement

# ── Couleurs CHG ───────────────────────────────────────────────────────────────
CHG_BLUE      = RGBColor(0x1F, 0x49, 0x7D)
WHITE         = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BLUE_BG = "DCE6F1"
CHG_BLUE_BG   = "1F497D"
GREEN_BG      = "E2EFDA"
RED_BG        = "FCE4D6"
ORANGE_BG     = "FFF2CC"
GREY_BG       = "F2F2F2"

MOIS_FR = {
    1:"janvier",2:"février",3:"mars",4:"avril",5:"mai",6:"juin",
    7:"juillet",8:"août",9:"septembre",10:"octobre",11:"novembre",12:"décembre"
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def set_cell_bg(cell, color_hex):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def set_table_border(table, color='1F497D'):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'), 'single')
        el.set(qn('w:sz'), '4')
        el.set(qn('w:color'), color)
        tblBorders.append(el)
    tblPr.append(tblBorders)

def set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)

def add_dropdown_cell(cell, choices, default, bg_hex="FFFFCC"):
    """Insert a Word native dropdown (w:sdt) into a table cell."""
    set_cell_bg(cell, bg_hex)
    nsmap = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    list_items = ""
    for c in choices:
        list_items += f'<w:listItem w:displayText="{c}" w:value="{c}"/>'
    sdt_xml = f"""<w:sdt {nsmap}>
  <w:sdtPr>
    <w:alias w:val="selection"/>
    <w:tag w:val="selection"/>
    <w:dropDownList>
      {list_items}
    </w:dropDownList>
  </w:sdtPr>
  <w:sdtContent>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="18"/>
        </w:rPr>
        <w:t>{default}</w:t>
      </w:r>
    </w:p>
  </w:sdtContent>
</w:sdt>"""
    sdt_element = etree.fromstring(sdt_xml)
    tc = cell._tc
    for child in list(tc):
        tc.remove(child)
    tc.append(sdt_element)

def no_space_para(para):
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)

def cell_para(cell, text, bold=False, size=9, color=None, bg=None, align=None, italic=False):
    if bg:
        set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    no_space_para(p)
    if align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(text) if text else "")
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        if isinstance(color, str):
            r, g, b = int(color[0:2],16), int(color[2:4],16), int(color[4:6],16)
            run.font.color.rgb = RGBColor(r, g, b)
        else:
            run.font.color.rgb = color

def add_section_title(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = CHG_BLUE
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bot = OxmlElement('w:bottom')
    bot.set(qn('w:val'), 'single')
    bot.set(qn('w:sz'), '6')
    bot.set(qn('w:color'), '1F497D')
    pBdr.append(bot)
    pPr.append(pBdr)

# ── Load libraries ─────────────────────────────────────────────────────────────

def load_cv_library():
    path = "/root/chg-library/cv_docx_library.json"
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        return list(data.values())
    return data

def load_projects_library():
    path = "/root/chg-library/projects_library.json"
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict) and 'projets' in data:
        return data['projets']
    if isinstance(data, list):
        return data
    return []

# ── Claude API call ────────────────────────────────────────────────────────────

def run_planning_analysis(gonogo_data, cvs, projects):
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    ident  = gonogo_data.get("identification", {})
    desc   = gonogo_data.get("description_projet", {})
    eval_  = gonogo_data.get("criteres_evaluation", {})
    ress   = gonogo_data.get("ressources_requises", {})
    strat  = gonogo_data.get("analyse_strategique", {})
    specs  = gonogo_data.get("specs_techniques_redaction", {})
    roles  = ress.get("roles_detailles", [])
    criteres = eval_.get("criteres_detail", [])

    # AO domain keywords for relevance filtering
    services_requis = desc.get("services_requis", [])
    titre_ao = ident.get("titre_projet", "")
    domaine_ao = f"{titre_ao} {' '.join(services_requis)} {desc.get('description','')}".lower()

    # Summarize CVs — include ALL CHG and DDM
    cv_summary = []
    for cv in cvs[:40]:
        nom    = cv.get("nom", "")
        texte  = cv.get("texte", "")[:600]
        source = cv.get("source", "CHG")
        cv_summary.append(f"NOM: {nom}\nSOURCE: {source}\nEXTRAIT: {texte}\n---")

    # Summarize projects
    proj_summary = []
    for p in projects[:41]:
        nom   = p.get("nom", "")
        texte = p.get("texte", "")[:400]
        proj_summary.append(f"FICHE: {nom}\nEXTRAIT: {texte}\n---")

    today = datetime.now().strftime("%d %B %Y")

    prompt = f"""Tu es un expert en planification d'offres de service pour une firme de génie-conseil québécoise (Groupe Conseil CHG).
DATE: {today}

CONTEXTE DE L'APPEL D'OFFRES:
- Titre: {ident.get('titre_projet','')}
- Client: {ident.get('client','')}
- Numéro AO: {ident.get('numero_ao','')}
- Localisation: {desc.get('localisation','')}
- Région: {desc.get('region_administrative','')}
- Description: {desc.get('description','')}
- Services requis: {', '.join(services_requis)}
- Date de dépôt: {ident.get('date_depot','')}

CRITÈRES D'ÉVALUATION DE L'AO:
{json.dumps(criteres, ensure_ascii=False, indent=2)}

RÔLES REQUIS PAR L'AO:
{json.dumps(roles, ensure_ascii=False, indent=2)}

POINTS FORTS CHG POUR CET AO:
{json.dumps(strat.get('points_forts_chg',[]), ensure_ascii=False)}

BIBLIOTHÈQUE CVs DISPONIBLES:
{"".join(cv_summary)}

BIBLIOTHÈQUE FICHES PROJETS DISPONIBLES:
{"".join(proj_summary)}

RÈGLES DE SÉLECTION DES RESSOURCES (par ordre de priorité):
1. PROXIMITÉ PHYSIQUE: prioriser les ressources dont le bureau est dans la même région que le projet
2. ANNÉES D'EXPÉRIENCE: si l'AO mentionne un nombre d'années minimum, prioriser les ressources senior. Si PAS mentionné, suggérer des profils juniors pour optimiser les coûts
3. EXPERTISE TECHNIQUE: correspondance avec les services requis par l'AO
4. Pour chaque rôle: suggérer 1 ressource PRINCIPALE + 1 ressource RELÈVE (backup)
5. Extraire les années d'expérience approximatives depuis le texte du CV (ex: "~15 ans" basé sur l'historique de carrière)

RÈGLES CRITIQUES POUR LES FICHES DE PROJETS:
- Sélectionner UNIQUEMENT des fiches dont le domaine correspond EXACTEMENT à l'AO
- AO de voirie/route → fiches de voirie, réfection de route, infrastructure routière UNIQUEMENT
- AO de structure → fiches de ponts, structures, bâtiments
- AO d'environnement → fiches environnementales
- EXCLURE toute fiche hors domaine (ex: si AO = route, exclure ponceaux isolés, bâtiments, foresterie)
- Pour chaque fiche suggérée: identifier quel chargé de projet de l'équipe proposée est le plus associé à cette fiche
- Maximum 8 fiches, toutes pertinentes au domaine exact de l'AO

Retourne UNIQUEMENT un JSON valide avec cette structure:
{{
  "criteres_succes": [
    {{
      "critere": "",
      "points": 0,
      "argumentaire_chg": "",
      "points_insistance": []
    }}
  ],
  "ressources_suggerees": [
    {{
      "role": "",
      "ressource_principale": "",
      "bureau_principal": "",
      "annees_experience_principale": "",
      "justification_principale": "",
      "ressource_releve": "",
      "bureau_releve": "",
      "annees_experience_releve": "",
      "justification_releve": "",
      "adequation_ao": ""
    }}
  ],
  "fiches_suggerees": [
    {{
      "nom_fiche": "",
      "pertinence": "",
      "similarites": "",
      "critere_couvert": "",
      "charge_projet_associe": ""
    }}
  ],
  "sous_traitance": [
    {{
      "discipline": "",
      "requis": false,
      "recommande": false,
      "raison": ""
    }}
  ],
  "notes_strategiques": ""
}}

Réponds UNIQUEMENT avec le JSON, sans texte avant ou après."""

    response = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}]
    )

    result_text = response.content[0].text.strip()
    if result_text.startswith("```"):
        result_text = result_text.split("\n", 1)[1].rsplit("```", 1)[0].strip()

    planning_data = json.loads(result_text)

    # ── Banque de ressources — ALL remaining CVs including DDM ────────────────
    assigned_names = set()
    for r in planning_data.get('ressources_suggerees', []):
        for field in ['ressource_principale', 'ressource_releve']:
            name = r.get(field, '').strip()
            if name:
                assigned_names.add(name.lower())

    banque_summary = []
    for cv in cvs:
        nom = cv.get('nom', '').strip()
        if nom.lower() in assigned_names:
            continue
        source = cv.get('source', 'CHG')
        texte = cv.get('texte', '')[:600]
        banque_summary.append(f"NOM: {nom}\nSOURCE: {source}\nEXTRAIT: {texte}\n----")

    region_projet  = desc.get('region_administrative', '') or ident.get('localisation', '')
    services_str   = ', '.join(services_requis)

    banque_prompt = (
        'Tu es un expert RH pour une firme de génie-conseil québécoise (Groupe Conseil CHG + DDM).\n\n'
        'CONTEXTE DU PROJET:\n'
        f'- Titre: {ident.get("titre_projet","")}\n'
        f'- Client: {ident.get("client","")}\n'
        f'- Région: {region_projet}\n'
        f'- Services requis: {services_str}\n\n'
        f'RESSOURCES DÉJÀ ASSIGNÉES (ne pas répéter):\n{list(assigned_names)}\n\n'
        'RESSOURCES DISPONIBLES (CHG + DDM):\n'
        + ''.join(banque_summary)
        + '\n\nRÈGLES:\n'
        '1. Exclure toutes les ressources déjà assignées\n'
        '2. Inclure TOUS les employés CHG et DDM restants pertinents pour ce type de projet\n'
        '3. Classer par: proximité bureau → années expérience → expertise technique\n'
        '4. DDM (environnement, foresterie, géomatique) = inclure si expertise complémentaire possible\n'
        '5. Maximum 15 ressources\n'
        '6. Inférer bureau (Québec/Montréal/Saguenay) depuis extrait CV\n'
        '7. Estimer années d\'expérience approximatives depuis historique de carrière dans le CV\n\n'
        'Retourne UNIQUEMENT un JSON valide avec clés plates:\n'
        '{"banque_count": N, '
        '"banque_0_nom": "Prenom Nom", '
        '"banque_0_role": "Ingénieur civil", '
        '"banque_0_bureau": "Québec", '
        '"banque_0_annees_experience": "~12 ans", '
        '"banque_0_justification": "Pourquoi pertinent", '
        '"banque_0_source": "CHG", '
        '"banque_1_nom": "..."}\n\n'
        'Réponds UNIQUEMENT avec le JSON, sans texte avant ou après.'
    )

    banque_data = {}
    try:
        banque_response = client.messages.create(
            model='claude-sonnet-4-5-20250929',
            max_tokens=4000,
            messages=[{'role': 'user', 'content': banque_prompt}]
        )
        banque_text = banque_response.content[0].text.strip()
        if banque_text.startswith('```'):
            banque_text = banque_text.split('\n', 1)[1].rsplit('```', 1)[0].strip()
        banque_data = json.loads(banque_text)
        print(f'  Banque: {banque_data.get("banque_count", 0)} ressources identifiées')
    except Exception as e:
        print(f'  Banque: erreur - {e}')
        banque_data = {'banque_count': 0}

    planning_data['banque_ressources'] = banque_data
    return planning_data

# ── Build Word document ────────────────────────────────────────────────────────

def build_planning_document(gonogo_data, planning_data, output_path):
    ident  = gonogo_data.get("identification", {})
    ech    = gonogo_data.get("echeancier_redaction", {})
    strat  = gonogo_data.get("analyse_strategique", {})
    specs  = gonogo_data.get("specs_techniques_redaction", {})

    criteres_succes = planning_data.get("criteres_succes", [])
    ressources      = planning_data.get("ressources_suggerees", [])
    fiches          = planning_data.get("fiches_suggerees", [])
    sous_traitance  = planning_data.get("sous_traitance", [])
    notes_strat     = planning_data.get("notes_strategiques", "")

    doc = Document()

    # Marges
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    # En-tête avec logo
    logo_path = "/root/chg-library/Logo_CHG.jpg"
    header = doc.sections[0].header
    htbl = header.add_table(1, 2, Cm(16))
    set_table_border(htbl, 'FFFFFF')
    hc0 = htbl.rows[0].cells[0]
    if os.path.exists(logo_path):
        hp = hc0.paragraphs[0]
        run = hp.add_run()
        run.add_picture(logo_path, width=Cm(4))
    hc1 = htbl.rows[0].cells[1]
    hp1 = hc1.paragraphs[0]
    hp1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp1.add_run("DOCUMENT DE PLANIFICATION")
    r.bold = True; r.font.size = Pt(10); r.font.color.rgb = CHG_BLUE
    now = datetime.now()
    hp1.add_run(f"\nGénéré le {now.day} {MOIS_FR[now.month]} {now.year}").font.size = Pt(8)

    # Titre
    titre = doc.add_paragraph()
    titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = titre.add_run(ident.get("titre_projet", "Document de planification").upper())
    tr.bold = True; tr.font.size = Pt(13); tr.font.color.rgb = CHG_BLUE
    titre.paragraph_format.space_before = Pt(6)
    titre.paragraph_format.space_after  = Pt(2)

    sous = doc.add_paragraph()
    sous.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sr = sous.add_run(f"{ident.get('client','')}  |  AO {ident.get('numero_ao','')}  |  Dépôt: {ident.get('date_depot','')}")
    sr.font.size = Pt(10); sr.font.color.rgb = CHG_BLUE
    sous.paragraph_format.space_after = Pt(10)

    # Info bloc rapide
    tbl_info = doc.add_table(rows=2, cols=4)
    tbl_info.style = 'Table Grid'
    set_table_border(tbl_info)
    labels = ["Score GO/NO-GO", "Recommandation", "Région", "Format dépôt"]
    values = [
        f"{strat.get('score_gonogo',0)}/100",
        strat.get('recommandation',''),
        gonogo_data.get('description_projet',{}).get('region_administrative',''),
        ident.get('format_depot','')
    ]
    for i in range(4):
        set_cell_bg(tbl_info.rows[0].cells[i], CHG_BLUE_BG)
        cell_para(tbl_info.rows[0].cells[i], labels[i], bold=True, color=WHITE, align='center', size=8)
        cell_para(tbl_info.rows[1].cells[i], values[i], bold=True, align='center', size=10)
    doc.add_paragraph()

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 1 — Spécifications techniques de rédaction
    # ══════════════════════════════════════════════════════════════════════════
    add_section_title(doc, "PARTIE 1 — SPÉCIFICATIONS TECHNIQUES DE RÉDACTION")

    tbl_specs = doc.add_table(rows=0, cols=2)
    tbl_specs.style = 'Table Grid'
    set_table_border(tbl_specs)
    set_col_width(tbl_specs, 0, 5.5)
    set_col_width(tbl_specs, 1, 11)

    def add_spec_row(table, label, value, highlight=False):
        row = table.add_row()
        set_cell_bg(row.cells[0], CHG_BLUE_BG if highlight else LIGHT_BLUE_BG)
        p0 = row.cells[0].paragraphs[0]
        no_space_para(p0)
        r0 = p0.add_run(label)
        r0.bold = True; r0.font.size = Pt(9)
        if highlight:
            r0.font.color.rgb = WHITE
        p1 = row.cells[1].paragraphs[0]
        no_space_para(p1)
        r1 = p1.add_run(str(value) if value else "Non spécifié")
        r1.font.size = Pt(9)
        if highlight and not value:
            r1.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
            r1.italic = True

    nb_pages = specs.get("nb_pages_max","")
    typo     = specs.get("typographie","")
    fmt      = specs.get("format_requis","")
    copies   = specs.get("nb_copies","")
    autres   = specs.get("autres_specs","")

    add_spec_row(tbl_specs, "Nombre de pages maximum", nb_pages, highlight=bool(nb_pages))
    add_spec_row(tbl_specs, "Typographie requise", typo, highlight=bool(typo))
    add_spec_row(tbl_specs, "Format requis", fmt)
    add_spec_row(tbl_specs, "Nombre de copies", copies)
    add_spec_row(tbl_specs, "Autres spécifications", autres)

    if not any([nb_pages, typo, fmt, copies, autres]):
        p_no_specs = doc.add_paragraph()
        r_no = p_no_specs.add_run("⚠ Aucune spécification technique détectée dans l'AO — vérifier le document source.")
        r_no.italic = True; r_no.font.size = Pt(9)
        r_no.font.color.rgb = RGBColor(0xC0, 0x70, 0x00)

    doc.add_paragraph()

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 2 — Critères de succès et argumentaire CHG
    # ══════════════════════════════════════════════════════════════════════════
    add_section_title(doc, "PARTIE 2 — CRITÈRES DE SUCCÈS ET ARGUMENTAIRE CHG")

    p_intro = doc.add_paragraph()
    r_intro = p_intro.add_run("Critères extraits de l'appel d'offres — points sur lesquels CHG doit insister dans sa réponse.")
    r_intro.italic = True; r_intro.font.size = Pt(9); r_intro.font.color.rgb = CHG_BLUE
    p_intro.paragraph_format.space_after = Pt(6)

    if criteres_succes:
        tbl2 = doc.add_table(rows=1, cols=3)
        tbl2.style = 'Table Grid'
        set_table_border(tbl2)
        for i, h in enumerate(["Critère de succès", "Points", "Argumentaire CHG — Points à insister"]):
            set_cell_bg(tbl2.rows[0].cells[i], CHG_BLUE_BG)
            cell_para(tbl2.rows[0].cells[i], h, bold=True, color=WHITE, align='center')
        set_col_width(tbl2, 0, 5)
        set_col_width(tbl2, 1, 1.5)
        set_col_width(tbl2, 2, 10)

        for crit in criteres_succes:
            row = tbl2.add_row()
            cell_para(row.cells[0], crit.get("critere",""), bold=True)
            cell_para(row.cells[1], str(crit.get("points","")), align='center', bold=True)
            cell = row.cells[2]
            p_arg = cell.paragraphs[0]
            no_space_para(p_arg)
            r_arg = p_arg.add_run(crit.get("argumentaire_chg",""))
            r_arg.font.size = Pt(9)
            r_arg.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            for pt in crit.get("points_insistance", []):
                p_pt = cell.add_paragraph()
                no_space_para(p_pt)
                r_pt = p_pt.add_run(f"• {pt}")
                r_pt.font.size = Pt(9)
                r_pt.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
    else:
        p = doc.add_paragraph()
        p.add_run("Aucun critère d'évaluation détecté dans l'AO — à compléter manuellement.").italic = True

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 3 — Suggestions de ressources
    # ══════════════════════════════════════════════════════════════════════════
    add_section_title(doc, "PARTIE 3 — SUGGESTIONS DE RESSOURCES")

    p_intro3 = doc.add_paragraph()
    r_intro3 = p_intro3.add_run("Ressources proposées selon: proximité physique → années d'expérience → expertise technique.")
    r_intro3.italic = True; r_intro3.font.size = Pt(9); r_intro3.font.color.rgb = CHG_BLUE
    p_intro3.paragraph_format.space_after = Pt(6)

    if ressources:
        tbl3 = doc.add_table(rows=1, cols=7)
        tbl3.style = 'Table Grid'
        set_table_border(tbl3)
        hdrs3 = ["Rôle", "Ressource principale", "Bureau", "Exp.", "Ressource relève", "Justification / Adéquation AO", "✓ Sélectionner"]
        for i, h in enumerate(hdrs3):
            set_cell_bg(tbl3.rows[0].cells[i], CHG_BLUE_BG)
            cell_para(tbl3.rows[0].cells[i], h, bold=True, color=WHITE, align='center')
        set_col_width(tbl3, 0, 2.5)
        set_col_width(tbl3, 1, 3)
        set_col_width(tbl3, 2, 1.5)
        set_col_width(tbl3, 3, 1.2)
        set_col_width(tbl3, 4, 2.5)
        set_col_width(tbl3, 5, 4)
        set_col_width(tbl3, 6, 2)

        for res in ressources:
            row = tbl3.add_row()
            cell_para(row.cells[0], res.get("role",""), bold=True)
            cell_para(row.cells[1], res.get("ressource_principale",""))
            cell_para(row.cells[2], res.get("bureau_principal",""), align='center')
            cell_para(row.cells[3], res.get("annees_experience_principale",""), align='center', size=8)
            releve_txt = res.get("ressource_releve","")
            bureau_r   = res.get("bureau_releve","")
            exp_r      = res.get("annees_experience_releve","")
            releve_full = releve_txt
            if bureau_r:
                releve_full += f"\n({bureau_r})"
            if exp_r:
                releve_full += f"\n{exp_r}"
            cell_para(row.cells[4], releve_full)
            just = res.get("justification_principale","")
            adeq = res.get("adequation_ao","")
            full = just
            if adeq and adeq != just:
                full += f"\n→ {adeq}"
            cell_para(row.cells[5], full)
            add_dropdown_cell(row.cells[6], ["Principale", "Relève", "Exclure"], "Principale")
    else:
        p = doc.add_paragraph()
        p.add_run("Aucune ressource suggérée — à compléter manuellement.").italic = True

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 4 — Banque de ressources disponibles
    # ══════════════════════════════════════════════════════════════════════════
    doc.add_paragraph()
    add_section_title(doc, "PARTIE 4 — BANQUE DE RESSOURCES DISPONIBLES")

    p_intro_b = doc.add_paragraph()
    r_intro_b = p_intro_b.add_run(
        "Ressources supplémentaires pertinentes non assignées en Partie 3. "
        "Sélectionnez Oui pour les ajouter à l'équipe. "
        "Classées par: proximité → expérience → expertise. "
        "Ressources DDM indiquées en orange."
    )
    r_intro_b.italic = True; r_intro_b.font.size = Pt(9); r_intro_b.font.color.rgb = CHG_BLUE
    p_intro_b.paragraph_format.space_after = Pt(6)

    banque_data  = planning_data.get('banque_ressources', {})
    banque_count = int(banque_data.get('banque_count', 0))
    banque_items = []
    for i in range(banque_count):
        item = {
            'nom':            banque_data.get(f'banque_{i}_nom', ''),
            'role':           banque_data.get(f'banque_{i}_role', ''),
            'bureau':         banque_data.get(f'banque_{i}_bureau', ''),
            'annees_exp':     banque_data.get(f'banque_{i}_annees_experience', ''),
            'justification':  banque_data.get(f'banque_{i}_justification', ''),
            'source':         banque_data.get(f'banque_{i}_source', 'CHG'),
        }
        if item['nom']:
            banque_items.append(item)

    if banque_items:
        tbl_b = doc.add_table(rows=1, cols=6)
        tbl_b.style = 'Table Grid'
        set_table_border(tbl_b)
        hdrs_b = ['Nom', 'Rôle', 'Bureau', 'Exp.', 'Justification pertinence', 'Sélectionner']
        for i, h in enumerate(hdrs_b):
            set_cell_bg(tbl_b.rows[0].cells[i], CHG_BLUE_BG)
            cell_para(tbl_b.rows[0].cells[i], h, bold=True, color=WHITE, align='center')
        set_col_width(tbl_b, 0, 3.5)
        set_col_width(tbl_b, 1, 3)
        set_col_width(tbl_b, 2, 1.8)
        set_col_width(tbl_b, 3, 1.2)
        set_col_width(tbl_b, 4, 5.5)
        set_col_width(tbl_b, 5, 1.5)

        for item in banque_items:
            row = tbl_b.add_row()
            is_ddm = item['source'].upper() == 'DDM'
            name_color = "C55A11" if is_ddm else None
            cell_para(row.cells[0], item['nom'], bold=True, color=name_color)
            cell_para(row.cells[1], item['role'])
            cell_para(row.cells[2], item['bureau'], align='center')
            cell_para(row.cells[3], item['annees_exp'], align='center', size=8)
            cell_para(row.cells[4], item['justification'])
            add_dropdown_cell(row.cells[5], ['Oui', 'Non'], 'Non')
    else:
        p_nb = doc.add_paragraph()
        p_nb.add_run('Aucune ressource supplémentaire identifiée.').italic = True

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 5 — Suggestions de fiches de projets comparables
    # ══════════════════════════════════════════════════════════════════════════
    add_section_title(doc, "PARTIE 5 — SUGGESTIONS DE FICHES DE PROJETS COMPARABLES")

    p_intro5 = doc.add_paragraph()
    r_intro5 = p_intro5.add_run(
        "Fiches suggérées comme projets comparables. Utilisez le menu déroulant pour indiquer "
        "où chaque fiche sera utilisée dans l'offre de service."
    )
    r_intro5.italic = True; r_intro5.font.size = Pt(9); r_intro5.font.color.rgb = CHG_BLUE
    p_intro5.paragraph_format.space_after = Pt(6)

    if fiches:
        tbl5 = doc.add_table(rows=1, cols=6)
        tbl5.style = 'Table Grid'
        set_table_border(tbl5)
        hdrs5 = ["Fiche de projet", "Pertinence", "Similarités avec l'AO", "Critère couvert", "Chargé de projet associé", "✓ Inclure"]
        for i, h in enumerate(hdrs5):
            set_cell_bg(tbl5.rows[0].cells[i], CHG_BLUE_BG)
            cell_para(tbl5.rows[0].cells[i], h, bold=True, color=WHITE, align='center')
        set_col_width(tbl5, 0, 3.5)
        set_col_width(tbl5, 1, 1.8)
        set_col_width(tbl5, 2, 4.5)
        set_col_width(tbl5, 3, 2)
        set_col_width(tbl5, 4, 2.5)
        set_col_width(tbl5, 5, 2.5)

        for fiche in fiches:
            row = tbl5.add_row()
            cell_para(row.cells[0], fiche.get("nom_fiche",""), bold=True)
            pertinence = fiche.get("pertinence","")
            bg_pert = (GREEN_BG if "élevée" in pertinence.lower() or "haute" in pertinence.lower()
                       else (ORANGE_BG if "moyenne" in pertinence.lower() else LIGHT_BLUE_BG))
            cell_para(row.cells[1], pertinence, bg=bg_pert, align='center')
            cell_para(row.cells[2], fiche.get("similarites",""))
            cell_para(row.cells[3], fiche.get("critere_couvert",""))
            cell_para(row.cells[4], fiche.get("charge_projet_associe",""), italic=True)
            # 4-choice dropdown
            add_dropdown_cell(row.cells[5],
                              ["Firme", "Chargé de projet", "Les deux", "Exclure"],
                              "Firme")
    else:
        p = doc.add_paragraph()
        p.add_run("Aucune fiche suggérée — à compléter manuellement.").italic = True

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 6 — Sous-traitance
    # ══════════════════════════════════════════════════════════════════════════
    sous_requis = [s for s in sous_traitance if s.get("requis") or s.get("recommande")]
    if sous_requis:
        add_section_title(doc, "PARTIE 6 — SOUS-TRAITANCE IDENTIFIÉE")
        tbl_st = doc.add_table(rows=1, cols=3)
        tbl_st.style = 'Table Grid'
        set_table_border(tbl_st)
        for i, h in enumerate(["Discipline", "Statut", "Raison"]):
            set_cell_bg(tbl_st.rows[0].cells[i], CHG_BLUE_BG)
            cell_para(tbl_st.rows[0].cells[i], h, bold=True, color=WHITE, align='center')
        set_col_width(tbl_st, 0, 4)
        set_col_width(tbl_st, 1, 2.5)
        set_col_width(tbl_st, 2, 10)
        for s in sous_requis:
            row = tbl_st.add_row()
            cell_para(row.cells[0], s.get("discipline",""), bold=True)
            statut = "REQUIS" if s.get("requis") else "RECOMMANDÉ"
            bg_s = RED_BG if s.get("requis") else ORANGE_BG
            cell_para(row.cells[1], statut, bold=True, bg=bg_s, align='center')
            cell_para(row.cells[2], s.get("raison",""))

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 7 — Échéancier de rédaction
    # ══════════════════════════════════════════════════════════════════════════
    add_section_title(doc, "PARTIE 7 — ÉCHÉANCIER DE RÉDACTION")

    tbl_ech = doc.add_table(rows=1, cols=2)
    tbl_ech.style = 'Table Grid'
    set_table_border(tbl_ech)
    for i, h in enumerate(["Étape", "Date cible"]):
        set_cell_bg(tbl_ech.rows[0].cells[i], CHG_BLUE_BG)
        cell_para(tbl_ech.rows[0].cells[i], h, bold=True, color=WHITE, align='center')
    set_col_width(tbl_ech, 0, 10)
    set_col_width(tbl_ech, 1, 6.5)

    # Order matches gonogo_report.py
    is_physical = gonogo_data.get("identification",{}).get("soumission_physique") or \
                  gonogo_data.get("identification",{}).get("delai_purolator_requis")
    etapes = [
        ("Réunion de planification",     ech.get("date_reunion_planification","")),
        ("Retour ingénieur responsable", ech.get("date_limite_ingenieur","")),
        ("Document final",               ech.get("date_document_final","")),
        ("Vérification ISO / qualité",   ech.get("date_iso","")),
        ("Livraison Purolator",          ech.get("date_livraison_purolator","") if is_physical else None),
        ("DATE DE DÉPÔT",                ech.get("date_depot","")),
    ]
    for label, date_val in etapes:
        if date_val is None:
            continue
        row = tbl_ech.add_row()
        is_depot = "DÉPÔT" in label
        if is_depot:
            set_cell_bg(row.cells[0], CHG_BLUE_BG)
            set_cell_bg(row.cells[1], CHG_BLUE_BG)
            cell_para(row.cells[0], label, bold=True, color=WHITE)
            cell_para(row.cells[1], date_val or "—", bold=True, color=WHITE, align='center')
        else:
            cell_para(row.cells[0], label)
            cell_para(row.cells[1], date_val or "À calculer", align='center')

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 8 — Instructions supplémentaires pour la rédaction
    # ══════════════════════════════════════════════════════════════════════════
    doc.add_paragraph()
    add_section_title(doc, "PARTIE 8 — INSTRUCTIONS SUPPLÉMENTAIRES POUR LA RÉDACTION")

    p_instruct_intro = doc.add_paragraph()
    r_ii = p_instruct_intro.add_run(
        "Ajoutez ici toute information contextuelle utile pour la rédaction de l'offre de service : "
        "relations avec le client, angles à privilégier, ressources à substituer, textes à injecter, "
        "contraintes particulières."
    )
    r_ii.italic = True; r_ii.font.size = Pt(9)
    r_ii.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    p_instruct_intro.paragraph_format.space_after = Pt(6)

    tbl_instruct = doc.add_table(rows=1, cols=1)
    tbl_instruct.style = 'Table Grid'
    set_table_border(tbl_instruct)
    instruct_cell = tbl_instruct.rows[0].cells[0]
    set_cell_bg(instruct_cell, "FAFAF0")
    p_instruct = instruct_cell.paragraphs[0]
    no_space_para(p_instruct)
    r_instruct = p_instruct.add_run(
        "[ Indiquez ici vos instructions : ex. Le client connaît déjà CHG via le mandat X — "
        "insister sur la continuité. Remplacer Jean-François Martin par Hugo Fontaine. "
        "Éviter de mentionner les délais passés sur le projet Y. ]"
    )
    r_instruct.italic = True; r_instruct.font.size = Pt(9)
    r_instruct.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    for _ in range(6):
        p_blank = instruct_cell.add_paragraph()
        no_space_para(p_blank)
        p_blank.add_run(" ")
    doc.add_paragraph()

    # ══════════════════════════════════════════════════════════════════════════
    # PARTIE 9 — Notes stratégiques
    # ══════════════════════════════════════════════════════════════════════════
    if notes_strat:
        add_section_title(doc, "PARTIE 9 — NOTES STRATÉGIQUES")
        items = re.split(r'(?<=[.!?])\s+(?=\d+[\.\)])|\n+|(?<=\w)\s{2,}', notes_strat.strip())
        items = [i.strip() for i in items if i.strip()]
        if len(items) <= 1:
            items = [i.strip() for i in notes_strat.split('.') if i.strip() and len(i.strip()) > 20]
        for idx, item in enumerate(items, 1):
            p_item = doc.add_paragraph()
            p_item.paragraph_format.left_indent = Pt(12)
            p_item.paragraph_format.space_before = Pt(2)
            r_num = p_item.add_run(f"{idx}. ")
            r_num.bold = True; r_num.font.size = Pt(9); r_num.font.color.rgb = CHG_BLUE
            r_text = p_item.add_run(item.lstrip('0123456789.-) '))
            r_text.italic = True; r_text.font.size = Pt(9)

    # Pied de page — validation
    doc.add_paragraph()
    p_conf = doc.add_paragraph()
    p_conf.paragraph_format.space_before = Pt(14)
    r_conf = p_conf.add_run(
        "VALIDATION — Ce document de planification doit être approuvé par le chargé de projet "
        "avant de procéder à la rédaction de l'offre de service. "
        "Les ressources et fiches suggérées peuvent être modifiées selon disponibilité et pertinence."
    )
    r_conf.italic = True; r_conf.font.size = Pt(8)
    r_conf.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    doc.save(output_path)
    print(f"Document de planification généré: {output_path}")
    return output_path


# ── Main ───────────────────────────────────────────────────────────────────────

def run_planning_agent(gonogo_json_path, output_path):
    print(f"Chargement Go/No-Go: {gonogo_json_path}")
    with open(gonogo_json_path, encoding="utf-8") as f:
        gonogo_data = json.load(f)

    print("Chargement bibliothèques CVs et projets...")
    cvs      = load_cv_library()
    projects = load_projects_library()
    print(f"  {len(cvs)} CVs | {len(projects)} fiches projets")

    print("Analyse Claude en cours...")
    planning_data = run_planning_analysis(gonogo_data, cvs, projects)

    print(f"  {len(planning_data.get('criteres_succes',[]))} critères de succès")
    print(f"  {len(planning_data.get('ressources_suggerees',[]))} ressources suggérées")
    print(f"  {len(planning_data.get('fiches_suggerees',[]))} fiches suggérées")

    print("Génération du document Word...")
    build_planning_document(gonogo_data, planning_data, output_path)

    json_out = output_path.replace(".docx", ".json")
    with open(json_out, "w", encoding="utf-8") as f:
        json.dump(planning_data, f, ensure_ascii=False, indent=2)
    print(f"JSON planification sauvegardé: {json_out}")

    return output_path


if __name__ == "__main__":
    import sys
    gonogo_path = sys.argv[1] if len(sys.argv) > 1 else "/root/test_ao_gonogo.json"
    output_path = sys.argv[2] if len(sys.argv) > 2 else "/local-files/PLANIFICATION_TEST.docx"
    run_planning_agent(gonogo_path, output_path)
