import os
import json
import anthropic
from docx import Document
from pathlib import Path

client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
OUTPUT = "/root/chg-library/real_proposal_analysis.json"

def extract_text(filepath):
    try:
        doc = Document(filepath)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return None

def analyze_os(filepath, index, total):
    name = Path(filepath).stem[:60]
    print(f"  [{index}/{total}] {name}")
    text = extract_text(filepath)
    if not text or len(text) < 500:
        print(f"    ⚠️ Trop court, ignoré")
        return None

    # Ask for plain text analysis, not JSON
    prompt = f"""Analyse cette offre de service réelle de CHG. 
Réponds en texte simple (PAS de JSON, PAS de guillemets spéciaux).

Extrais:
TON: [décris le ton et style en 2-3 phrases]
STRUCTURE: [liste les sections utilisées]
FORMULES: [cite 3-5 formules d'introduction ou conclusion typiques]
ARGUMENTS: [liste 3-5 arguments différenciateurs]
VOCABULAIRE: [liste 10-15 mots techniques clés]

OS:
{text[:5000]}"""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}]
        )
        return {
            "fichier": name,
            "analyse": response.content[0].text
        }
    except Exception as e:
        print(f"    ⚠️ Erreur: {e}")
        return None

def main():
    print("Apprentissage OS — SharePoint v2")
    print("=" * 50)

    all_files = []
    for root, dirs, files in os.walk("/root/chg-sharepoint/Travaux en cours"):
        if "OLD" in root:
            continue
        for f in files:
            if f.endswith("_ar.docx") and not f.startswith("~"):
                all_files.append(os.path.join(root, f))

    # Add textes généraux
    textes_dir = "/root/chg-sharepoint/CHG Conseil/Offres de service/Textes généraux"
    for f in os.listdir(textes_dir):
        if f.endswith(".docx") and not f.startswith("~"):
            all_files.append(os.path.join(textes_dir, f))

    print(f"\n📄 {len(all_files)} fichiers — analyse des 15 meilleurs\n")

    analyses = []
    for i, fpath in enumerate(all_files[:15], 1):
        result = analyze_os(fpath, i, 15)
        if result:
            analyses.append(result)
            print(f"    ✅ OK")

    # Generate synthesis
    print(f"\n🧠 Génération de la synthèse ({len(analyses)} OS analysées)...")
    combined = "\n\n---\n\n".join([f"FICHIER: {a['fichier']}\n{a['analyse']}" for a in analyses])
    
    synth_prompt = f"""Tu es un expert en rédaction d'offres de service en génie-conseil québécois.
À partir de ces analyses de vraies OS de CHG, génère un guide de style complet.

{combined[:8000]}

Génère un guide structuré avec:
1. STYLE GLOBAL: ton, registre, niveau de langue
2. FORMULES TYPES: phrases d'introduction et conclusion à réutiliser
3. STRUCTURE RECOMMANDÉE: ordre des sections
4. ARGUMENTS GAGNANTS: ce qui différencie CHG
5. VOCABULAIRE CLÉ: termes techniques à utiliser
6. RÈGLES DE FRANÇAIS: points d'orthographe et grammaire spécifiques au domaine

Réponds en texte structuré, clair, directement utilisable."""

    response = client.messages.create(
        model="claude-sonnet-4-5-20250929",
        max_tokens=3000,
        messages=[{"role": "user", "content": synth_prompt}]
    )
    synthese = response.content[0].text

    output = {
        "nb_fichiers_analyses": len(analyses),
        "analyses": analyses,
        "synthese": synthese
    }

    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(output, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Analyse complète sauvegardée: {OUTPUT}")
    print(f"📊 {len(analyses)}/15 OS analysées avec succès")
    print("\n=== SYNTHÈSE ===")
    print(synthese[:500])

if __name__ == "__main__":
    main()
