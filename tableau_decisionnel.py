#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Session 23 — Tableau decisionnel GO-NOGO (updated per Alexandra's feedback)
Populates the 10 "blue" columns of Alexandra's template with data from a
Go/No-Go JSON, leaving the 4 "gray" columns (11-14) completely blank.
Formats multi-item content as line-broken bullets (no commas).

Session 23 changes:
  - Col 3 (Qualité/Prix): now surfaces explicit mode_evaluation classification
  - Col 8 (Ressources): filters out AMP authorization mentions (irrelevant here);
    merges "Responsable de projet" / "Chargé de projet" into one role (same function)
  - Col 9 (Fiches de projet): split into "Firme" vs "Responsable de projet"
    groups, each showing nb projets requis + délai (années) explicitly
  - Col 10 (Assurances): removed duplicate montant line (redundant with
    assurances_requises bullets)

Usage:
    python3 tableau_decisionnel.py <template.xlsx> <gonogo.json> <pdf_path> <output.xlsx>

Note (Session 23): calling convention changed — pdf_path is now required
(3rd argument) so the bordereau de prix can be located and rendered from
the source AO PDF. Update any caller (e.g. pipeline_w1.py) accordingly.
"""
import sys
import json
import re
import os
import tempfile
import openpyxl
from copy import copy
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from bordereau_extractor import render_bordereau_image
import uuid
import subprocess
import shutil


def bullets(value):
    """
    Format a value as line-broken bullet points.
    - list -> one bullet per item
    - string with commas -> split into bullets on each comma
    - plain string / empty -> returned as-is (or empty string)
    """
    if value is None:
        return ""
    if isinstance(value, list):
        items = [str(v).strip() for v in value if str(v).strip()]
        if not items:
            return ""
        return "\n".join(f"• {v}" for v in items)
    if isinstance(value, str):
        v = value.strip()
        if not v:
            return ""
        if "," in v:
            parts = [p.strip() for p in v.split(",") if p.strip()]
            if len(parts) > 1:
                return "\n".join(f"• {p}" for p in parts)
        return v
    return str(value)


def combine(*parts):
    """Join non-empty bulleted blocks with a blank line between groups."""
    blocks = [p for p in parts if p]
    return "\n".join(blocks)


PLACEHOLDER_VALUES = {"non spécifié", "non specifie", "n/a", "non applicable", "", "-", "—"}


def is_placeholder(value):
    if value is None:
        return True
    return str(value).strip().lower() in PLACEHOLDER_VALUES


def dedup_lists(*lists):
    """
    Merge multiple lists of strings into one, removing case-insensitive /
    near-duplicate entries while preserving first-seen order and casing.
    """
    seen = set()
    merged = []
    for lst in lists:
        for item in lst or []:
            item = str(item).strip()
            if not item:
                continue
            key = item.lower()
            if key in seen:
                continue
            seen.add(key)
            merged.append(item)
    return merged


def normalize_words(text):
    """Lowercase, strip punctuation, return set of words for similarity comparison."""
    t = re.sub(r'[^\w\s]', ' ', str(text).lower())
    return set(t.split())


def is_similar_text(a, b, threshold=0.5):
    """
    True if two strings share enough words to be considered the same
    underlying requirement, even with different phrasing (e.g. 'Membre en
    règle de l'Ordre des ingénieurs du Québec' vs 'Membre de l'Ordre des
    ingénieurs du Québec (responsable de projet)').
    """
    wa, wb = normalize_words(a), normalize_words(b)
    if not wa or not wb:
        return False
    overlap = wa & wb
    return len(overlap) / min(len(wa), len(wb)) >= threshold


def split_echeancier_milestones(raw_text):
    """
    Splits a raw echeancier_client string into individual milestone
    bullets. Hierarchical strategy, in priority order:
      1. Split on sentence-ending periods (". " or final ".") first —
         respects genuine sentence/phase boundaries.
      2. Within each resulting chunk, split on ";" if present.
      3. Within each remaining piece, split on "," but NEVER on a
         comma that falls inside an open "(" ... ")" pair — this is
         the fix for cases like "Relevés (2 sem, avril 2026)" where a
         duration and a date share one parenthetical.
    Returns a list of cleaned, non-empty milestone strings.

    Session 49 (Alexandra, via Pat): validated against 10 real
    echeancier_client samples spanning comma-lists, parenthetical-comma
    dates, semicolon-lists, and mixed period+semicolon phase text. Known
    remaining gap, out of scope for this patch: pipe ("|") delimited
    sources (e.g. AO_DATA_batch_1785167324) aren't split by any tier
    here and render as a single unbroken bullet — future item.
    """
    def paren_aware_comma_split(s):
        parts = []
        depth = 0
        current = []
        for ch in s:
            if ch == '(':
                depth += 1
                current.append(ch)
            elif ch == ')':
                depth = max(0, depth - 1)
                current.append(ch)
            elif ch == ',' and depth == 0:
                parts.append(''.join(current))
                current = []
            else:
                current.append(ch)
        parts.append(''.join(current))
        return parts

    milestones = []
    sentence_chunks = re.split(r'\.(?:\s+|$)', raw_text)

    for chunk in sentence_chunks:
        chunk = chunk.strip()
        if not chunk:
            continue

        semi_parts = [p.strip() for p in chunk.split(';') if p.strip()]

        for part in semi_parts:
            comma_parts = paren_aware_comma_split(part)
            for cp in comma_parts:
                cp = cp.strip().rstrip('.').strip()
                if cp:
                    milestones.append(cp)

    return milestones


def dedup_services_livrables(services, livrables):
    """
    Merge services_requis and livrables into one list for Column 5,
    removing both exact duplicates (dedup_lists) and near-duplicates
    where a livrable restates a services_requis entry in different
    wording (e.g. "Rapport de conception" vs "Préparation du rapport
    de conception, plans et devis préliminaires...").

    Session 49 (Alexandra, via Pat): near-duplicate matching is scoped
    to services-vs-livrables comparisons only — never within the same
    list. An earlier broad version (any item vs any item) was stress-
    tested against every real AO in production data and produced false
    positives that silently dropped genuinely distinct services (e.g.
    "Volet RTU et éclairage" and "Volet environnement" both collapsing
    into an unrelated "Volet mobilité..." entry, purely on the shared
    word "Volet"). Restricting to cross-list comparison eliminated every
    observed false positive while still catching all real duplicate
    pairs in the data.
    """
    merged = dedup_lists(services, livrables)
    services_lower = {str(s).strip().lower() for s in services or []}
    kept_services = [item for item in merged if item.strip().lower() in services_lower]

    result = []
    for item in merged:
        if item.strip().lower() in services_lower:
            result.append(item)
        elif not any(is_similar_text(item, s) for s in kept_services):
            result.append(item)
    return result


def dedup_expertises(exp_resume, exp_cles):
    """
    Merge expertises_requises_resume and expertises_cles into one list for
    Column 6, removing both exact duplicates (dedup_lists) and near-
    duplicates where one list restates the other in different wording
    (e.g. "Ingénierie civile municipale" vs "Ingénierie civile").

    Same cross-list-only restriction as dedup_services_livrables() (Session
    49): near-duplicate matching only compares exp_cles entries against
    exp_resume entries, never within the same list, to avoid the false-
    positive collisions found when stress-testing the broader approach.
    """
    merged = dedup_lists(exp_resume, exp_cles)
    resume_lower = {str(s).strip().lower() for s in exp_resume or []}
    kept_resume = [item for item in merged if item.strip().lower() in resume_lower]

    result = []
    for item in merged:
        if item.strip().lower() in resume_lower:
            result.append(item)
        elif not any(is_similar_text(item, s) for s in kept_resume):
            result.append(item)
    return result


def format_expertise_categories(categories):
    """
    Session 51: format expertise_categories (from analyze_ao.py's
    validate_expertise_categories()) for Column E display. Canonical
    category names are passed through unchanged; "autre: X" / "autre : X"
    entries (lowercase, from the extraction schema) are relabeled to
    "Autre : X" to match the capitalization of the canonical bullets.
    """
    result = []
    for item in categories or []:
        text = str(item).strip()
        if not text:
            continue
        m = re.match(r"^autre\s*:\s*(.*)$", text, re.IGNORECASE)
        if m:
            result.append(f"Autre : {m.group(1).strip()}")
        else:
            result.append(text)
    return result


def yn(flag):
    if flag is True:
        return "Oui"
    if flag is False:
        return "Non"
    return ""

def oui_non_pas_specifie(val):
    # Session 47 fix: sous_traitance_autorisee / consortium_autorise are now
    # 3-state strings ("oui" / "non" / "non_specifie") instead of booleans,
    # so "not mentioned in the AO" can be distinguished from "explicitly
    # forbidden". Also tolerates legacy boolean values from older cached JSON.
    if val is True:
        return "Oui"
    if val is False:
        return "Non"
    if isinstance(val, str):
        v = val.strip().lower()
        if v == "oui":
            return "Oui"
        if v == "non":
            return "Non"
        if v == "non_specifie":
            return "Pas spécifié"
    return ""


def set_cell(ws, row, col, text):
    c = ws.cell(row=row, column=col)
    c.value = text
    existing = c.alignment
    c.alignment = openpyxl.styles.Alignment(
        wrap_text=True,
        vertical=existing.vertical or "top",
        horizontal=existing.horizontal,
    )


# ── Session 23: AMP filter (item #4) ────────────────────────────────────────
def is_amp_entry(text):
    """True if text refers to the AMP (Autorité des marchés publics) authorization."""
    t = str(text).lower()
    if "autorité des marchés publics" in t:
        return True
    if re.search(r'\bamp\b', t):
        return True
    return False


# ── Session 23: role merge (item #7) ────────────────────────────────────────
ROLE_SYNONYMS = {
    "responsable de projet": "responsable_chef",
    "chargé de projet": "responsable_chef",
    "charge de projet": "responsable_chef",
    "chef de projet": "responsable_chef",
}


def merge_roles(roles):
    """
    Merge 'Responsable de projet' / 'Chargé de projet' / 'Chef de projet'
    into a single canonical role entry (same function at CHG per Alexandra).
    Keeps the richer data from whichever source has it: max years experience,
    union of qualifications, non-placeholder title, non-empty points_differenciateurs.
    """
    grouped = {}
    order = []
    for r in roles:
        role_name = (r.get("role") or "").strip().lower()
        key = ROLE_SYNONYMS.get(role_name)
        if key is None:
            order.append(dict(r))
            continue
        if key not in grouped:
            entry = dict(r)
            entry["role"] = "Responsable de projet"
            grouped[key] = entry
            order.append(entry)
        else:
            existing = grouped[key]
            existing["annees_experience_min"] = max(
                existing.get("annees_experience_min") or 0,
                r.get("annees_experience_min") or 0,
            )
            existing["annees_experience_mentionnees"] = (
                existing.get("annees_experience_mentionnees") or r.get("annees_experience_mentionnees")
            )
            existing["qualifications_obligatoires"] = dedup_lists(
                existing.get("qualifications_obligatoires", []),
                r.get("qualifications_obligatoires", []),
            )
            existing["qualifications_souhaitees"] = dedup_lists(
                existing.get("qualifications_souhaitees", []),
                r.get("qualifications_souhaitees", []),
            )
            if is_placeholder(existing.get("titre_requis")) and not is_placeholder(r.get("titre_requis")):
                existing["titre_requis"] = r.get("titre_requis")
            if not existing.get("points_differenciateurs") and r.get("points_differenciateurs"):
                existing["points_differenciateurs"] = r.get("points_differenciateurs")
    return order


# ── Session 23: fiches classification (item #5) ─────────────────────────────
def classify_fiche_target(intitule):
    """Classify a criteres_detail entry as requiring fiches for 'firme' or 'responsable'."""
    t = (intitule or "").lower()
    # Session 32 fix: match singular OR plural forms (e.g. "responsables de projet")
    # so multi-role AOs don't silently fall through to the "firme" default.
    # Session 49 fix: "responsable DU projet" (contraction) was not matched
    # by the "de" form, silently falling through to the "firme" default.
    if re.search(r"responsables?\s+d[eu]\s+projet", t) or \
       re.search(r"charg(é|és)\s+de\s+projet", t) or \
       re.search(r"chefs?\s+de\s+projet", t):
        return "responsable"
    if "prestataire" in t or "firme" in t or "soumissionnaire" in t or "entreprise" in t:
        return "firme"
    return None


def main():
    if len(sys.argv) != 5:
        print("Usage: python3 tableau_decisionnel.py <template.xlsx> <gonogo.json> <pdf_path> <output.xlsx>")
        sys.exit(1)

    template_path, json_path, pdf_path, output_path = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]

    # Session 36: pdf_path may be a JSON file listing multiple candidate
    # paths (multi-fichiers batch mode). If it doesn't parse as a JSON list,
    # treat it as a single real PDF path — identical to pre-Session-36 behavior.
    _bordereau_candidates = [pdf_path]
    try:
        with open(pdf_path, "r", encoding="utf-8") as _bcf:
            _parsed_candidates = json.load(_bcf)
        if isinstance(_parsed_candidates, list) and _parsed_candidates:
            _bordereau_candidates = _parsed_candidates
    except Exception:
        pass

    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    ident = data.get("identification", {}) or {}
    descr = data.get("description_projet", {}) or {}
    conf = data.get("criteres_conformite", {}) or {}
    eval_ = data.get("criteres_evaluation", {}) or {}
    ress = data.get("ressources_requises", {}) or {}
    risq = data.get("risques_contractuels", {}) or {}
    strat = data.get("analyse_strategique", {}) or {}

    wb = openpyxl.load_workbook(template_path)
    ws = wb["Feuil1"] if "Feuil1" in wb.sheetnames else wb.active

    ROW = 2  # single-row demo output for now

    # --- Column 1: Titre | Numero AO | N. SEAO ---
    titre = ident.get("titre_projet", "")
    numero = ident.get("numero_ao", "")
    col1 = combine(
        f"Titre : {titre}" if titre else "",
        f"N° AO : {numero}" if numero else "",
    )
    set_cell(ws, ROW, 1, col1)

    # --- Column 2: Date de depot | Format ---
    date_depot = ident.get("date_depot", "")
    heure_depot = ident.get("heure_depot", "")
    format_depot = ident.get("format_depot", "")
    purolator = ident.get("delai_purolator_requis", False)
    lines2 = []
    if date_depot:
        lines2.append(f"• Date limite : {date_depot}" + (f" {heure_depot}" if heure_depot else ""))
    if format_depot:
        lines2.append(f"• Format : {format_depot}")
    if purolator:
        lines2.append("• Soumission physique — délai Purolator (+2 j. ouvrables) à prévoir")
    set_cell(ws, ROW, 2, "\n".join(lines2))

    # --- Column 3: Qualite/prix (Session 24: simplified to mode_evaluation only, per Alexandra) ---
    mode_eval = eval_.get("mode_evaluation", "")
    MODE_EVAL_DISPLAY = {
        "Prix seul": "Prix",
        "Qualité seule": "Qualité",
        "Qualité et prix": "Qualité et prix",
    }
    mode_eval_display = MODE_EVAL_DISPLAY.get(mode_eval, mode_eval)
    lines3 = []
    if mode_eval_display:
        lines3.append(f"Mode d'évaluation : {mode_eval_display}")
    set_cell(ws, ROW, 3, "\n".join(lines3))

    # --- Column 4: Description du projet ---
    sommaire = descr.get("sommaire_projet", "") or descr.get("description", "")
    localisation = descr.get("localisation", "")
    lines4 = []
    if sommaire:
        lines4.append(sommaire)
    if localisation:
        lines4.append(f"• Localisation : {localisation}")
    set_cell(ws, ROW, 4, "\n".join(lines4))

    # --- Column 5: Services a fournir (Session 51: reads expertise_categories
    # instead of services_requis/livrables -- see classify_expertise_categories()
    # in analyze_ao.py) ---
    expertise_categories = descr.get("expertise_categories", [])
    col5 = bullets(format_expertise_categories(expertise_categories))
    set_cell(ws, ROW, 5, col5)

    # --- Column 6: Expertises requises ---
    exp_resume = descr.get("expertises_requises_resume", [])
    exp_cles = ress.get("expertises_cles", [])
    exp_merged = dedup_expertises(exp_resume, exp_cles)
    set_cell(ws, ROW, 6, bullets(exp_merged))

    # --- Column 7: Sous-traitant/Consortium ---
    # Session 49 (Alexandra Robitaille, via Pat): stripped to exactly these
    # two lines. sous_traitance_recommandee / domaines_sous_traitance are no
    # longer rendered here (kept in the JSON/extraction, not shown anywhere).
    st_autorisee = ress.get("sous_traitance_autorisee", None)
    consortium = risq.get("consortium_autorise", None)
    lines7 = []
    if st_autorisee is not None:
        lines7.append(f"• Sous-traitance autorisée : {oui_non_pas_specifie(st_autorisee)}")
    if consortium is not None:
        lines7.append(f"• Consortium autorisé : {oui_non_pas_specifie(consortium)}")
    set_cell(ws, ROW, 7, "\n".join([l for l in lines7 if l]))

    # --- Column 8: Ressources et qualifications minimales (Session 23: AMP filtered, roles merged, empty roles dropped, certifs deduped) ---
    roles = merge_roles(ress.get("roles_detailles", []))
    lines8 = []
    seen_quals = []
    for r in roles:
        role = r.get("role", "")
        annees = r.get("annees_experience_min", "")
        quals_obl = r.get("qualifications_obligatoires", [])
        if not role or (not annees and not quals_obl):
            continue
        entry = f"• {role}"
        if annees:
            entry += f" — {annees} ans exp. min."
        lines8.append(entry)
        for q in quals_obl:
            lines8.append(f"    - {q}")
            seen_quals.append(q)
    certifs = [
        c for c in conf.get("formations_certifications_requises", [])
        if not is_amp_entry(c) and not any(is_similar_text(c, q) for q in seen_quals)
    ]
    if certifs:
        lines8.append(bullets(certifs))
    set_cell(ws, ROW, 8, "\n".join([l for l in lines8 if l]))

    # --- Column 9: Fiches de projet (Session 23: split Firme vs Responsable de projet) ---
    criteres_detail = eval_.get("criteres_detail", [])
    budget = descr.get("budget_estime", "")
    montant_min = ress.get("montant_minimal_projets", "")

    firme_lines = []
    resp_lines = []
    autres_criteres = []  # Session 49 fix: (intitule, entry) for criteria that
    # don't match "firme"/"prestataire" or a "responsable de projet" pattern --
    # previously silently folded into firme_lines, which produced unlabeled,
    # colliding numbers under "Firme" for criteria like "autres membres du
    # personnel" or "surveillant de chantier" that have nothing to do with the
    # firm's own track record.
    for c in criteres_detail:
        nb_proj = c.get("nb_projets_requis", "")
        annees_exp = c.get("annees_experience_requises", "")
        if not (nb_proj or annees_exp):
            continue
        entry = "• "
        if nb_proj:
            entry += f"{nb_proj} projet(s) comparable(s) requis"
        else:
            entry += "Projets comparables requis"
        # Session 32 fix: the fiches/projects realization window ("réalisés au
        # cours des X dernières années") is a distinct concept from
        # annees_experience_requises (years of experience required for the
        # role itself) and was previously mislabeled using the latter, which
        # produced wrong numbers whenever the two figures differ (e.g. "10 ans
        # d'expérience" but fiches "réalisés au cours des 7 dernières années").
        # No separate structured field exists for this yet, so extract it from
        # the contenu_requis free text instead (confirmed max one match per
        # critère across sampled AOs).
        # Session 47 fix: AOs sometimes spell the number out with the digit
        # in parentheses ("dix (10) dernières années") instead of a bare
        # digit ("10 dernières années"); the old regex only matched the
        # latter and silently dropped the window in the former case.
        m = re.search(
            r"\(?(\d+)\)?\s+derni[eè]res?\s+ann[eé]es",
            c.get("contenu_requis", ""),
            re.IGNORECASE,
        )
        if m:
            entry += f" au cours des {m.group(1)} dernières années"
        target = classify_fiche_target(c.get("intitule"))
        if target == "responsable":
            resp_lines.append(entry)
        elif target == "firme":
            firme_lines.append(entry)
        else:
            autres_criteres.append((c.get("intitule") or "Autre critère", entry))

    lines9 = []
    if firme_lines:
        lines9.append("Fiches requises — Firme :")
        lines9.extend(firme_lines)
    if resp_lines:
        if lines9:
            lines9.append("")
        lines9.append("Fiches requises — Responsable de projet :")
        lines9.extend(resp_lines)
    for intitule, entry in autres_criteres:
        if lines9:
            lines9.append("")
        lines9.append(f"Fiches requises — {intitule} :")
        lines9.append(entry)
    # Session 53 (Karine feedback, AO 307109): budget/montant_min/sous_traitance/
    # consortium are AO-level fields, never reliably tied to one specific
    # criterion -- confirmed via the prompt's own rules (document-wide search,
    # not criterion-scoped) and via real AO samples where montant_minimal_projets
    # sometimes spans multiple criteria in one string. Previously appended
    # directly after the per-criterion blocks with no separation, which
    # visually (and misleadingly) attached them to whichever criterion
    # happened to render last -- in 307109's case, "Qualité des rapports"
    # (27.4), even though the $3M figure's own citation says "critère 27.1".
    # Given their own clearly labeled section instead of being reattached to
    # any specific criterion, which doesn't generalize (see Session 53
    # investigation across 4 real AOs). Also restores sous_traitance_autorisee
    # here -- previously read only for Column 7, never for this cell, despite
    # this cell's own label promising it.
    infos_generales = []
    if not is_placeholder(budget):
        infos_generales.append(f"• Budget estimé (AO) : {budget}")
    if not is_placeholder(montant_min):
        infos_generales.append(f"• Valeur minimale projets de référence (AO) : {montant_min}")
    if st_autorisee is not None:
        infos_generales.append(f"• Sous-traitance permise (AO) : {oui_non_pas_specifie(st_autorisee)}")
    if consortium is not None:
        infos_generales.append(f"• Consortium permis (AO) : {oui_non_pas_specifie(consortium)}")
    if infos_generales:
        if lines9:
            lines9.append("")
        lines9.append("Informations générales (AO) :")
        lines9.extend(infos_generales)
    set_cell(ws, ROW, 9, "\n".join(lines9))

    # --- Column 10: Assurances requises (et montant) (Session 23: amount attached per-line when missing) ---
    assurances = conf.get("assurances_requises", [])
    montant_assur = conf.get("montant_assurance_responsabilite", "")
    has_dollar = any("$" in str(a) for a in assurances)
    montant_by_type = {}
    if not has_dollar and montant_assur:
        for amt, typ in re.findall(r'([\d\s,\.]+\$)[^()]*\(([^)]+)\)', montant_assur):
            montant_by_type[typ.strip().lower()] = amt.strip()

    lines10 = []
    for a in assurances:
        line = f"• {a}"
        if not has_dollar:
            matched_amt = None
            for typ, amt in montant_by_type.items():
                type_words = [w for w in re.split(r'\s+', typ) if w.lower() not in ("rc", "et", "de", "la") and len(w) > 3]
                if any(w.lower() in str(a).lower() for w in type_words):
                    matched_amt = amt
                    break
            if matched_amt:
                line += f" — Montant : {matched_amt}"
        lines10.append(line)
    # Fallback: if amounts exist but couldn't be matched to a specific line
    # (no type parsed / no keyword match), still surface the raw montant text
    if not has_dollar and montant_assur and not montant_by_type:
        lines10.append(f"• Montant : {montant_assur}")
    set_cell(ws, ROW, 10, "\n".join([l for l in lines10 if l]))

    # --- Column 11 (NEW, Session 23, item #8): Bordereau de prix — thumbnail + link ---
    # Inserted after "Assurances requises" — shifts existing gray columns 11-14
    # (Questions/validations, Décision, Responsable rédaction, Chargé de projet)
    # to 12-15. Blue "populate" styling copied from column 10's header cell.
    # Session 23 revision: full-size image was too large for a comfortable
    # cell — now saves the full image to a persistent, HTTP-served path and
    # embeds a small thumbnail + clickable hyperlink instead (Pat's feedback).
    BORDEREAU_COL = 11

    ws.insert_cols(idx=BORDEREAU_COL, amount=1)

    header_source = ws.cell(row=1, column=10)  # "Assurances requises" header, pre-shift styling reference
    header_new = ws.cell(row=1, column=BORDEREAU_COL)
    header_new.value = "Bordereau de prix (aperçu)"
    if header_source.has_style:
        header_new.font = copy(header_source.font)
        header_new.fill = copy(header_source.fill)
        header_new.border = copy(header_source.border)
        header_new.alignment = copy(header_source.alignment)

    col_letter = get_column_letter(BORDEREAU_COL)
    ws.column_dimensions[col_letter].width = 26

    safe_numero = re.sub(r'[^A-Za-z0-9_-]', '_', str(ident.get("numero_ao") or "AO"))
    # Session 38 fix: numero_ao extraction is not guaranteed unique or even
    # correct per-AO (confirmed collision: APM and Lotbinière both wrote to
    # Bordereau_26-780.png). Append a run-unique token derived from
    # output_path (the batch_<ts>_<hash> segment present in every pipeline
    # output filename) so concurrent/sequential runs can never collide,
    # while keeping numero_ao for human readability.
    _run_token_match = re.search(r'(batch_[A-Za-z0-9_]+)', os.path.basename(output_path))
    run_token = _run_token_match.group(1) if _run_token_match else uuid.uuid4().hex[:8]
    persistent_png = f"/local-files/Bordereau_{safe_numero}_{run_token}.png"
    # Note: the VPS's http.server (port 8888) runs with cwd=/local-files, so
    # it serves that directory AS the URL root — no /local-files/ segment in
    # the URL itself (confirmed via readlink /proc/<pid>/cwd, Session 23).
    public_url = f"http://177.7.34.82:8888/Bordereau_{safe_numero}_{run_token}.png"

    _XLSX_EXTS = (".xlsx", ".xls", ".xlsm")
    try:
        result_path, pages = None, []
        _bordereau_attempted = 0   # candidates where a real scan was actually attempted (excludes unsupported-type skips)
        _bordereau_clean = 0       # candidates where render_bordereau_image() completed without raising
        _bordereau_errors = 0      # candidates where xlsx conversion or the scan itself raised/failed
        for _cand_path in _bordereau_candidates:
            _cand_lower = str(_cand_path).lower()
            _pdf_for_render = None
            _temp_outdir = None
            _temp_profile_dir = None

            if _cand_lower.endswith(".pdf"):
                _pdf_for_render = _cand_path
                _bordereau_attempted += 1
            elif _cand_lower.endswith(_XLSX_EXTS):
                _bordereau_attempted += 1
                # Session 38 (Shawinigan case): bordereaux can arrive as real
                # .xlsx files with no PDF equivalent. Convert to a temporary
                # PDF via headless LibreOffice, then reuse the existing
                # render_bordereau_image() unchanged — its detection is
                # format-agnostic once the input is a PDF (anchors on
                # heading text + currency evidence via text extraction).
                # Isolated profile/output dirs per conversion prevent the
                # same class of concurrent-run collision fixed earlier this
                # session for thumbnail filenames.
                try:
                    _temp_profile_dir = tempfile.mkdtemp(prefix="soffice_profile_")
                    _temp_outdir = tempfile.mkdtemp(prefix="xlsx2pdf_")
                    _conv_result = subprocess.run(
                        [
                            "soffice", "--headless", "--convert-to", "pdf",
                            "--outdir", _temp_outdir,
                            f"-env:UserInstallation=file://{_temp_profile_dir}",
                            _cand_path,
                        ],
                        capture_output=True, text=True, timeout=45,
                    )
                    _base = os.path.splitext(os.path.basename(_cand_path))[0]
                    _candidate_pdf = os.path.join(_temp_outdir, f"{_base}.pdf")
                    if _conv_result.returncode == 0 and os.path.exists(_candidate_pdf):
                        _pdf_for_render = _candidate_pdf
                        print(f"  [bordereau] converted xlsx to PDF: {_cand_path}")
                    else:
                        _bordereau_errors += 1
                        print(f"  [bordereau] xlsx conversion failed for {_cand_path}: rc={_conv_result.returncode} stderr={_conv_result.stderr[-300:]}")
                except subprocess.TimeoutExpired:
                    _bordereau_errors += 1
                    print(f"  [bordereau] xlsx conversion timed out: {_cand_path}")
                except Exception as _conv_e:
                    _bordereau_errors += 1
                    print(f"  [bordereau] xlsx conversion error for {_cand_path}: {_conv_e}")
            else:
                print(f"  [bordereau] skipping unsupported candidate type: {_cand_path}")
                continue

            try:
                if _pdf_for_render is None:
                    continue
                _rp, _pg = render_bordereau_image(_pdf_for_render, persistent_png)
                _bordereau_clean += 1
                if _rp:
                    result_path, pages = _rp, _pg
                    print(f"  [bordereau] match found in: {_cand_path}")
                    break
            except Exception as _cand_e:
                _bordereau_errors += 1
                print(f"  [bordereau] error scanning {_cand_path}: {_cand_e}")
                continue
            finally:
                if _temp_outdir and os.path.isdir(_temp_outdir):
                    shutil.rmtree(_temp_outdir, ignore_errors=True)
                if _temp_profile_dir and os.path.isdir(_temp_profile_dir):
                    shutil.rmtree(_temp_profile_dir, ignore_errors=True)
        if result_path:
            from PIL import Image as PILImage
            from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
            from openpyxl.drawing.xdr import XDRPositiveSize2D
            from openpyxl.utils.units import pixels_to_EMU

            thumb_path = os.path.join(tempfile.gettempdir(), f"bordereau_thumb_{os.getpid()}.png")
            with PILImage.open(persistent_png) as full_img:
                thumb_width = 160
                ratio = thumb_width / full_img.width
                thumb_height = int(full_img.height * ratio)
                full_img.resize((thumb_width, thumb_height)).save(thumb_path)

            pages_str = ", ".join(f"p.{p}" for p in pages)
            link_cell = ws.cell(row=ROW, column=BORDEREAU_COL)
            link_cell.value = f"🔗 Voir le bordereau complet ({pages_str})"
            link_cell.hyperlink = public_url
            link_cell.font = openpyxl.styles.Font(color="0563C1", underline="single", size=10)
            link_cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")

            img = XLImage(thumb_path)
            text_reserve_px = 45  # space left at top of cell for the link text
            marker = AnchorMarker(col=BORDEREAU_COL - 1, colOff=pixels_to_EMU(5), row=ROW - 1, rowOff=pixels_to_EMU(text_reserve_px))
            img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(thumb_width), pixels_to_EMU(thumb_height)))
            ws.add_image(img)

            needed_height_pts = (text_reserve_px + thumb_height + 10) * 0.75
            ws.row_dimensions[ROW].height = max(ws.row_dimensions[ROW].height or 0, needed_height_pts)
            print(f"Bordereau saved to {persistent_png} (page(s) {pages}) — thumbnail {thumb_width}x{thumb_height}px, link: {public_url}")
        else:
            if _bordereau_clean == 0 and _bordereau_errors > 0:
                set_cell(ws, ROW, BORDEREAU_COL, "Erreur lors de l'extraction du bordereau — à vérifier manuellement.")
                print(f"WARNING: bordereau detection could not complete — {_bordereau_errors} of {_bordereau_attempted} candidate(s) errored, 0 clean scans")
            else:
                set_cell(ws, ROW, BORDEREAU_COL, "Bordereau de prix non détecté dans le document source.")
                if _bordereau_errors > 0:
                    print(f"WARNING: no bordereau page detected — placeholder text set instead ({_bordereau_errors} of {_bordereau_attempted} candidate(s) errored, {_bordereau_clean} clean scan(s) found nothing)")
                else:
                    print("WARNING: no bordereau page detected — placeholder text set instead")
    except Exception as e:
        set_cell(ws, ROW, BORDEREAU_COL, "Erreur lors de l'extraction du bordereau — à vérifier manuellement.")
        print(f"WARNING: bordereau extraction failed: {e}")

    # --- Column 12 (NEW, Session 29, Alexandra feedback #1): Échéancier ---
    ECHEANCIER_COL = 12
    ws.insert_cols(idx=ECHEANCIER_COL, amount=1)
    header_source_ech = ws.cell(row=1, column=BORDEREAU_COL)
    header_ech = ws.cell(row=1, column=ECHEANCIER_COL)
    header_ech.value = "Échéancier"
    if header_source_ech.has_style:
        header_ech.font = copy(header_source_ech.font)
        header_ech.fill = copy(header_source_ech.fill)
        header_ech.border = copy(header_source_ech.border)
        header_ech.alignment = copy(header_source_ech.alignment)
    ws.column_dimensions[get_column_letter(ECHEANCIER_COL)].width = 26

    echeancier_client = descr.get("echeancier_client", "")
    if echeancier_client:
        etapes = split_echeancier_milestones(echeancier_client)
        set_cell(ws, ROW, ECHEANCIER_COL, "\n\n".join(f"• {e}" for e in etapes))
    else:
        set_cell(ws, ROW, ECHEANCIER_COL, "")

    # --- Column 13 (NEW, Session 29, Alexandra feedback #4): Intrants fournis par le client ---
    INTRANTS_COL = 13
    ws.insert_cols(idx=INTRANTS_COL, amount=1)
    header_source_int = ws.cell(row=1, column=ECHEANCIER_COL)
    header_int = ws.cell(row=1, column=INTRANTS_COL)
    header_int.value = "Intrants fournis par le client"
    if header_source_int.has_style:
        header_int.font = copy(header_source_int.font)
        header_int.fill = copy(header_source_int.fill)
        header_int.border = copy(header_source_int.border)
        header_int.alignment = copy(header_source_int.alignment)
    ws.column_dimensions[get_column_letter(INTRANTS_COL)].width = 30

    intrants = descr.get("intrants_fournis_client", [])
    set_cell(ws, ROW, INTRANTS_COL, bullets(intrants))

    # --- Column 14 (NEW, Session 53, Karine feedback item 5): Visite des lieux obligatoire ---
    VISITE_LIEUX_COL = 14
    ws.insert_cols(idx=VISITE_LIEUX_COL, amount=1)
    header_source_visite = ws.cell(row=1, column=INTRANTS_COL)
    header_visite = ws.cell(row=1, column=VISITE_LIEUX_COL)
    header_visite.value = "Visite des lieux obligatoire"
    if header_source_visite.has_style:
        header_visite.font = copy(header_source_visite.font)
        header_visite.fill = copy(header_source_visite.fill)
        header_visite.border = copy(header_source_visite.border)
        header_visite.alignment = copy(header_source_visite.alignment)
    ws.column_dimensions[get_column_letter(VISITE_LIEUX_COL)].width = 22

    set_cell(ws, ROW, VISITE_LIEUX_COL, oui_non_pas_specifie(conf.get("visite_lieux_obligatoire")))

    # --- Column 15 (NEW, Session 54): Cumul de rôles autorisé ---
    CUMUL_ROLES_COL = 15
    ws.insert_cols(idx=CUMUL_ROLES_COL, amount=1)
    header_source_cumul = ws.cell(row=1, column=VISITE_LIEUX_COL)
    header_cumul = ws.cell(row=1, column=CUMUL_ROLES_COL)
    header_cumul.value = "Cumul de rôles autorisé"
    if header_source_cumul.has_style:
        header_cumul.font = copy(header_source_cumul.font)
        header_cumul.fill = copy(header_source_cumul.fill)
        header_cumul.border = copy(header_source_cumul.border)
        header_cumul.alignment = copy(header_source_cumul.alignment)
    ws.column_dimensions[get_column_letter(CUMUL_ROLES_COL)].width = 32

    set_cell(ws, ROW, CUMUL_ROLES_COL, oui_non_pas_specifie(ress.get("cumul_roles_autorise")))

    # --- Column 16 (NEW, Session 54): Autres disciplines procurées séparément ---
    AUTRES_DISC_COL = 16
    ws.insert_cols(idx=AUTRES_DISC_COL, amount=1)
    header_source_disc = ws.cell(row=1, column=CUMUL_ROLES_COL)
    header_disc = ws.cell(row=1, column=AUTRES_DISC_COL)
    header_disc.value = "Autres disciplines procurées séparément"
    if header_source_disc.has_style:
        header_disc.font = copy(header_source_disc.font)
        header_disc.fill = copy(header_source_disc.fill)
        header_disc.border = copy(header_source_disc.border)
        header_disc.alignment = copy(header_source_disc.alignment)
    ws.column_dimensions[get_column_letter(AUTRES_DISC_COL)].width = 32

    set_cell(ws, ROW, AUTRES_DISC_COL, oui_non_pas_specifie(risq.get("autres_disciplines_separees")))

    # --- Row height floor for readability of other columns ---
    if not ws.row_dimensions[ROW].height or ws.row_dimensions[ROW].height < 260:
        ws.row_dimensions[ROW].height = 260

    # Columns 17-20 (shifted from original 11-14) intentionally left untouched (blank, gray)

    wb.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
