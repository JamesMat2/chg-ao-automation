import anthropic
import json
import os
import re
from datetime import date

# ── Text extraction ────────────────────────────────────────────────────────────

def extract_text_pdf(path):
    try:
        from pdfminer.high_level import extract_text
        text = extract_text(path)
        return text[:50000] if text else ""
    except Exception as e:
        print(f"Erreur extraction PDF: {e}")
        return ""

def extract_text_docx(path):
    try:
        from docx import Document
        doc = Document(path)
        parts = []
        for para in doc.paragraphs:
            if para.text.strip():
                parts.append(para.text.strip())
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        parts.append(cell.text.strip())
        return "\n".join(parts)[:50000]
    except Exception as e:
        print(f"Erreur extraction DOCX: {e}")
        return ""

def extract_ao_text(path):
    if not path:
        return ""
    ext = os.path.splitext(path)[1].lower()
    if ext in ['.docx', '.doc']:
        text = extract_text_docx(path)
    else:
        text = extract_text_pdf(path)
    return text

# ── Post-processor: normalize criteres_detail keys ────────────────────────────

def normalize_criteres_detail(criteres):
    """
    Claude sometimes returns variant key names for criteres_detail.
    This normalizes all variants to the exact keys the renderer expects:
      numero_critere, intitule, points_max, contenu_requis,
      nb_projets_requis, annees_experience_requises
    """
    normalized = []
    for c in criteres:
        n = {}
        # numero_critere
        n["numero_critere"] = (
            c.get("numero_critere") or
            c.get("numero") or
            c.get("critere") or
            c.get("id") or ""
        )
        # intitule
        n["intitule"] = (
            c.get("intitule") or
            c.get("titre") or
            c.get("nom") or
            c.get("name") or ""
        )
        # points_max
        raw_pts = (
            c.get("points_max") or
            c.get("points") or
            c.get("ponderation") or
            c.get("weight") or 0
        )
        try:
            n["points_max"] = int(raw_pts)
        except (ValueError, TypeError):
            n["points_max"] = 0
        # contenu_requis
        n["contenu_requis"] = (
            c.get("contenu_requis") or
            c.get("description") or
            c.get("contenu") or
            c.get("details") or
            c.get("exigences") or ""
        )
        # nb_projets_requis
        try:
            n["nb_projets_requis"] = int(c.get("nb_projets_requis") or c.get("nb_projets") or 0)
        except (ValueError, TypeError):
            n["nb_projets_requis"] = 0
        # annees_experience_requises
        try:
            n["annees_experience_requises"] = int(
                c.get("annees_experience_requises") or
                c.get("annees_experience") or
                c.get("experience_years") or 0
            )
        except (ValueError, TypeError):
            n["annees_experience_requises"] = 0

        normalized.append(n)
    return normalized

# ── Main analysis ──────────────────────────────────────────────────────────────

def analyze_ao_gonogo(pdf_path):
    """Analyse un AO et retourne la grille Go/No-Go complète"""

    print(f"Extraction du texte: {pdf_path}")
    ao_text = extract_ao_text(pdf_path)

    if not ao_text:
        raise Exception("Impossible d'extraire le texte du document")

    print(f"Texte extrait: {len(ao_text)} caractères")
    print("Analyse Claude en cours...")

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    today = date.today().strftime("%d %B %Y")

    # Jours fériés Québec 2025-2026
    feries_qc = """
JOURS FÉRIÉS QUÉBEC 2025-2026 (à exclure de l'échéancier):
- 2025-12-25 Noël
- 2025-12-26 Lendemain de Noël
- 2026-01-01 Jour de l'An
- 2026-01-02 Lendemain du Jour de l'An
- 2026-04-03 Vendredi saint
- 2026-05-18 Journée nationale des patriotes
- 2026-06-24 Fête nationale du Québec
- 2026-07-01 Fête du Canada
- 2026-09-07 Fête du Travail
- 2026-10-12 Action de grâces
- 2026-11-11 Jour du Souvenir (non férié QC mais noter)
- 2026-12-25 Noël
RÈGLE: Exclure TOUS les samedis, dimanches et jours fériés ci-dessus du calcul des dates.
RÈGLE PUROLATOR: Si format_depot = "physique" ou "papier" ou "en personne", ajouter +2 jours ouvrables supplémentaires pour la livraison Purolator.
"""

    prompt = f"""Tu es un expert en appels d'offres pour une firme de génie-conseil québécoise (Groupe Conseil CHG).
DATE D'AUJOURD'HUI: {today} — utilise cette date pour évaluer les délais. Si la date de dépôt est déjà passée, indique clairement que l'AO est EXPIRÉ.

{feries_qc}

DOCUMENT AO:
{ao_text}

Retourne UNIQUEMENT un JSON valide avec cette structure exacte — respecte TOUS les noms de clés exactement tels qu'écrits:
{{
  "identification": {{
    "numero_ao": "",
    "titre_projet": "",
    "client": "",
    "contact_client": "",
    "email_contact": "",
    "telephone_contact": "",
    "date_depot": "",
    "heure_depot": "",
    "format_depot": "",
    "lieu_depot": "",
    "date_limite_questions": "",
    "date_emission": "",
    "timezone_projet": "",
    "soumission_physique": false,
    "delai_purolator_requis": false
  }},
  "description_projet": {{
    "description": "",
    "sommaire_projet": "Résumé clair en 3-4 phrases du projet: quoi, où, pourquoi, ampleur. Langage non-technique compréhensible par tous.",
    "expertises_requises_resume": ["expertise 1", "expertise 2"],
    "localisation": "",
    "region_administrative": "",
    "services_requis": [],
    "livrables": [],
    "echeancier_client": "",
    "budget_estime": ""
  }},
  "specs_techniques_redaction": {{
    "nb_pages_max": "",
    "typographie": "",
    "format_requis": "",
    "nb_copies": "",
    "autres_specs": ""
  }},
  "criteres_conformite": {{
    "visite_obligatoire": false,
    "date_visite": "",
    "garantie_soumission_requise": false,
    "montant_garantie": "",
    "assurances_requises": [],
    "montant_assurance_responsabilite": "",
    "logiciels_obligatoires": [],
    "formations_certifications_requises": [],
    "criteres_eliminatoires": []
  }},
  "criteres_evaluation": {{
    "mode_attribution": "",
    "ponderation_technique": "",
    "ponderation_financiere": "",
    "note_minimale_qualite": "",
    "entrevue_prevue": false,
    "organigramme_requis": false,
    "nombre_repondants_attendus": "",
    "criteres_detail": [
      {{
        "numero_critere": "2.7.1",
        "intitule": "Nom du critère",
        "points_max": 20,
        "contenu_requis": "Description détaillée de ce qui est évalué",
        "nb_projets_requis": 5,
        "annees_experience_requises": 10
      }}
    ]
  }},
  "ressources_requises": {{
    "expertises_cles": [],
    "cv_requis": true,
    "fiches_projets_requises": true,
    "nombre_cv": "",
    "montant_minimal_projets": "",
    "mode_remuneration": "",
    "sous_traitance_autorisee": false,
    "sous_traitance_recommandee": false,
    "domaines_sous_traitance": [],
    "roles_detailles": [
      {{
        "role": "",
        "titre_requis": "",
        "annees_experience_min": 0,
        "annees_experience_mentionnees": false,
        "qualifications_obligatoires": [],
        "qualifications_souhaitees": [],
        "points_differenciateurs": "",
        "profil_junior_acceptable": false,
        "nb_personnes": 1
      }}
    ]
  }},
  "incongruites": [
    {{
      "section": "",
      "description": "",
      "question_a_poser": ""
    }}
  ],
  "risques_contractuels": {{
    "type_contrat": "",
    "penalites": "",
    "clauses_responsabilite": "",
    "consortium_autorise": false,
    "autres_risques": []
  }},
  "analyse_strategique": {{
    "historique_client_chg": "",
    "niveau_concurrence_estime": "",
    "concurrents_region": [],
    "points_forts_chg": [],
    "points_faibles_chg": [],
    "score_gonogo": 0,
    "recommandation": "",
    "justification": ""
  }},
  "echeancier_redaction": {{
    "date_depot": "",
    "date_document_final": "",
    "date_limite_ingenieur": "",
    "date_iso": "",
    "date_reunion_planification": "",
    "date_livraison_purolator": "",
    "notes_echeancier": ""
  }}
}}

RÈGLES IMPORTANTES:
- Score Go/No-Go: 0-100 (70+ = GO, 50-69 = CONDITIONNEL, moins de 50 = NO GO)
- Si AO expiré: score = 0, recommandation = "NO GO", justification = "APPEL D'OFFRES EXPIRÉ"
- Détecte TOUTES les incongruités, contradictions ou ambiguïtés dans le document

RÈGLES CRITIQUES POUR criteres_detail — RESPECTER EXACTEMENT CES NOMS DE CLÉS:
  * "numero_critere" (pas "numero", pas "critere", pas "id")
  * "intitule" (pas "titre", pas "nom", pas "name")
  * "points_max" (pas "points", pas "ponderation", pas "weight")
  * "contenu_requis" (pas "description", pas "contenu", pas "details")
  * "nb_projets_requis" (nombre entier, 0 si non spécifié)
  * "annees_experience_requises" (nombre entier, 0 si non spécifié)

RÈGLES POUR roles_detailles:
- Extraire CHAQUE rôle mentionné dans l'AO avec ses exigences exactes
- Si annees_experience non mentionnées: annees_experience_mentionnees = false, profil_junior_acceptable = true
- "points_differenciateurs": identifier si ce rôle génère des points selon un critère d'évaluation.
  Format: "Critère 2.7.X — X pts — [condition qui donne des points ex: 10 ans exp = X pts]"
  Si aucun lien direct avec un critère: laisser vide ""

RÈGLES POUR specs_techniques_redaction:
- Extraire le nombre de pages maximum si mentionné dans l'AO
- Extraire la typographie requise (police, taille) si mentionnée
- Extraire le nombre de copies requises
- Si non spécifié: laisser la valeur vide ""

RÈGLES POUR sommaire_projet:
- 3 à 4 phrases maximum
- Langage clair, non-technique
- Inclure: quoi (type de travaux), où (localisation), ampleur (budget/distance/superficie), contexte si pertinent

RÈGLES POUR concurrents_region:
- Laisser [] — sera rempli via analyse historique CHG (Phase 2)

RÈGLES POUR l'échéancier (calcul à rebours, weekends ET jours fériés exclus):
  * date_reunion_planification: dès que possible après émission
  * date_document_final: 4 jours ouvrables avant dépôt (2e étape après réunion)
  * date_limite_ingenieur: 6 jours ouvrables avant dépôt
  * date_iso: 2 jours ouvrables avant dépôt
  * date_livraison_purolator: si soumission physique, date_depot - 2 jours ouvrables
  * Si soumission physique: intégrer +2 jours ouvrables pour Purolator dans tous les calculs

- Pour timezone_projet: noter si le projet est dans un fuseau différent de EST
- Réponds UNIQUEMENT avec le JSON, sans texte avant ou après, sans backticks markdown"""

    response = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=8000,
        messages=[{"role": "user", "content": prompt}]
    )

    result_text = response.content[0].text.strip()

    # Strip markdown fences if present
    if result_text.startswith("```"):
        result_text = result_text.split("\n", 1)[1].rsplit("```", 1)[0].strip()

    result = json.loads(result_text)

    # ── Post-processor 1: normalize criteres_detail keys (schema enforcement) ──
    try:
        criteres_raw = result.get("criteres_evaluation", {}).get("criteres_detail", [])
        if criteres_raw:
            result["criteres_evaluation"]["criteres_detail"] = normalize_criteres_detail(criteres_raw)
            print(f"  criteres_detail normalisés: {len(result['criteres_evaluation']['criteres_detail'])} critères")
    except Exception as e:
        print(f"  [criteres normalization] avertissement: {e}")

    # ── Post-processor 2: expiry detection (Python is authoritative, not Claude) ──
    try:
        from datetime import date as _date
        import re as _re

        _MOIS_FR = {
            "janvier":1,"février":2,"fevrier":2,"mars":3,"avril":4,
            "mai":5,"juin":6,"juillet":7,"août":8,"aout":8,
            "septembre":9,"octobre":10,"novembre":11,"décembre":12,"decembre":12
        }
        _MOIS_EN = {
            "january":1,"february":2,"march":3,"april":4,"may":5,"june":6,
            "july":7,"august":8,"september":9,"october":10,"november":11,"december":12
        }

        def _parse_date(s):
            """Parse FR / EN / ISO date strings. Returns date or None."""
            if not s:
                return None
            s = s.strip()
            # ISO: 2025-11-25
            m = _re.match(r'(\d{4})-(\d{2})-(\d{2})', s)
            if m:
                return _date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            # French: 25 novembre 2025
            m = _re.match(r'(\d{1,2})\s+([a-zA-Zéûôùàâè]+)\s+(\d{4})', s, _re.IGNORECASE)
            if m:
                mois = m.group(2).lower()
                num  = _MOIS_FR.get(mois) or _MOIS_EN.get(mois)
                if num:
                    return _date(int(m.group(3)), num, int(m.group(1)))
            # English: November 25, 2025
            m = _re.match(r'([a-zA-Z]+)\s+(\d{1,2})[,\s]+(\d{4})', s, _re.IGNORECASE)
            if m:
                num = _MOIS_EN.get(m.group(1).lower())
                if num:
                    return _date(int(m.group(3)), num, int(m.group(2)))
            return None

        _date_depot_str = result.get("identification", {}).get("date_depot", "")
        _parsed = _parse_date(_date_depot_str)

        if _parsed and _parsed < _date.today():
            _strat = result.get("analyse_strategique", {})
            # Preserve analytical score and recommendation for reference
            _strat["score_analytique"]         = _strat.get("score_gonogo", 0)
            _strat["recommandation_analytique"] = _strat.get("recommandation", "—")
            # Override to EXPIRÉ
            _strat["score_gonogo"]    = 0
            _strat["recommandation"]  = "EXPIRÉ"
            _strat["justification"]   = (
                f"APPEL D'OFFRES EXPIRÉ — La date limite de dépôt était le "
                f"{_date_depot_str}. Cet appel d'offres ne peut plus être soumissionné. "
                f"Score analytique conservé à titre de référence : "
                f"{_strat['score_analytique']}/100 — {_strat['recommandation_analytique']}."
            )
            result["analyse_strategique"] = _strat
    except Exception as _e:
        print(f"  [expiry check] avertissement: {_e}")
    # ── Fin post-processors ──────────────────────────────────────────────────

    print("Analyse Go/No-Go complétée!")
    return result


if __name__ == "__main__":
    import sys
    pdf_path = sys.argv[1] if len(sys.argv) > 1 else "/root/test_ao.pdf"

    result = analyze_ao_gonogo(pdf_path)

    output_path = pdf_path.rsplit(".", 1)[0] + "_gonogo.json"
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\nRésultat sauvegardé: {output_path}")
    print(f"\nScore Go/No-Go: {result['analyse_strategique']['score_gonogo']}/100")
    print(f"Recommandation: {result['analyse_strategique']['recommandation']}")
    print(f"Incongruités détectées: {len(result['incongruites'])}")
    print(f"Rôles détaillés: {len(result['ressources_requises']['roles_detailles'])}")
    criteres = result.get('criteres_evaluation', {}).get('criteres_detail', [])
    print(f"Critères d'évaluation: {len(criteres)}")
    specs = result.get('specs_techniques_redaction', {})
    if any(specs.values()):
        print(f"Specs rédaction: pages max={specs.get('nb_pages_max','—')}, typo={specs.get('typographie','—')}")
