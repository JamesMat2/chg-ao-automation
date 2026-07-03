import os
import json
from docx import Document
from pathlib import Path

CV_DIRS = [
    "/root/chg-sharepoint/CHG Conseil/Offres de service/Descriptions du personnel",
    "/root/chg-sharepoint/CHG Conseil/Offres de service/Descriptions du personnel/À RÉVISER",
]

OUTPUT = "/root/chg-library/cv_docx_library.json"

def extract_cv(filepath):
    try:
        doc = Document(filepath)
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
        name = Path(filepath).stem
        name = name.replace(" (à réviser)", "").replace(" – Chargé de projet", "").replace(" – Chargée de projet", "").replace(" – Techn assigné au projet", "").strip()
        return {
            "nom": name,
            "fichier": os.path.basename(filepath),
            "texte": text[:8000],
            "source": "sharepoint"
        }
    except Exception as e:
        print(f"  ⚠️ Erreur {filepath}: {e}")
        return None

def main():
    print("Reconstruction de la bibliothèque CV")
    print("=" * 50)
    
    cvs = []
    seen = set()
    
    for cv_dir in CV_DIRS:
        if not os.path.exists(cv_dir):
            print(f"⚠️ Dossier non trouvé: {cv_dir}")
            continue
        files = [f for f in os.listdir(cv_dir) if f.endswith(".docx") and not f.startswith("~")]
        print(f"\n📁 {cv_dir}")
        print(f"   {len(files)} fichiers trouvés")
        for fname in sorted(files):
            fpath = os.path.join(cv_dir, fname)
            cv = extract_cv(fpath)
            if cv and cv["nom"] not in seen:
                seen.add(cv["nom"])
                cvs.append(cv)
                print(f"  ✅ {cv['nom']}")

    os.makedirs("/root/chg-library", exist_ok=True)
    with open(OUTPUT, "w", encoding="utf-8") as f:
        json.dump(cvs, f, ensure_ascii=False, indent=2)

    print(f"\n✅ {len(cvs)} CVs sauvegardés dans {OUTPUT}")

if __name__ == "__main__":
    main()
