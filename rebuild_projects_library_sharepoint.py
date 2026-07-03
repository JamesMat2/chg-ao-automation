import os
import json
from docx import Document
from pathlib import Path

PROJECTS_DIR = "/root/chg-sharepoint/CHG Conseil/Offres de service/Fiches de projet"
OUTPUT = "/root/chg-library/projects_library.json"

def extract_project(filepath):
    try:
        doc = Document(filepath)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        name = Path(filepath).stem
        return {
            "nom": name,
            "fichier": os.path.basename(filepath),
            "texte": text[:6000],
            "source": "sharepoint"
        }
    except Exception as e:
        print(f"  ⚠️ Erreur {filepath}: {e}")
        return None

def main():
    print("Reconstruction de la bibliothèque Fiches de projet")
    print("=" * 50)

    projects = []
    files = [f for f in os.listdir(PROJECTS_DIR) if f.endswith(".docx") and not f.startswith("~")]
    print(f"{len(files)} fiches trouvées\n")

    for fname in sorted(files):
        fpath = os.path.join(PROJECTS_DIR, fname)
        proj = extract_project(fpath)
        if proj:
            projects.append(proj)
            print(f"  ✅ {proj['nom'][:80]}")

    os.makedirs("/root/chg-library", exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(projects, f, ensure_ascii=False, indent=2)

    print(f"\n✅ {len(projects)} fiches sauvegardées dans {OUTPUT}")

if __name__ == "__main__":
    main()
