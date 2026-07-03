import os
import json
import requests
import msal
from pathlib import Path

CLIENT_ID     = os.environ["AZURE_CLIENT_ID"]
TENANT_ID     = os.environ["AZURE_TENANT_ID"]
CLIENT_SECRET = os.environ["AZURE_CLIENT_SECRET"]
AUTHORITY     = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES        = ["https://graph.microsoft.com/.default"]
SYNC_ROOT     = "/root/chg-sharepoint"
ALEXANDRA_UPN = "alexandra.robitaille@chgconseil.com"
TARGET_FOLDER = "CHG Automatisation AI"

def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" in result:
        print("✅ Authentifié avec succès (Graph API)")
        return result["access_token"]
    raise Exception(f"Erreur auth: {result.get('error_description')}")

def graph_get(token, url):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()

def download_file(token, url, local_path):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers, stream=True)
    r.raise_for_status()
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)

def sync_folder(token, drive_id, item_id, local_path, depth=0):
    os.makedirs(local_path, exist_ok=True)
    indent = "  " * depth
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children"
    data = graph_get(token, url)
    items = data.get("value", [])
    for item in items:
        name = item["name"]
        item_local = os.path.join(local_path, name)
        if "folder" in item:
            print(f"{indent}📁 {name}/ ({item['folder']['childCount']} items)")
            sync_folder(token, drive_id, item["id"], item_local, depth+1)
        elif "file" in item:
            size = item.get("size", 0)
            download_url = item.get("@microsoft.graph.downloadUrl")
            if download_url:
                print(f"{indent}📄 {name} ({size:,} bytes)")
                download_file(token, download_url, item_local)

def main():
    print("CHG Automatisation AI — Sync OneDrive")
    print("=" * 60)

    token = get_token()

    # Get Alexandra's drive
    print(f"\n📂 Connexion au drive de {ALEXANDRA_UPN}...")
    drive_data = graph_get(token, f"https://graph.microsoft.com/v1.0/users/{ALEXANDRA_UPN}/drive")
    drive_id = drive_data["id"]
    print(f"✅ Drive ID: {drive_id[:20]}...")

    # Find CHG Automatisation AI folder
    print(f"\n🔍 Recherche du dossier '{TARGET_FOLDER}'...")
    root_data = graph_get(token, f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children")
    
    chg_folder = None
    for item in root_data.get("value", []):
        if TARGET_FOLDER in item["name"]:
            chg_folder = item
            break

    if not chg_folder:
        print(f"❌ Dossier '{TARGET_FOLDER}' non trouvé")
        print("Dossiers disponibles:")
        for item in root_data.get("value", []):
            print(f"  - {item['name']}")
        return

    print(f"✅ Dossier trouvé: {chg_folder['name']}")
    print(f"\n📥 Début du sync vers {SYNC_ROOT}...\n")
    
    os.makedirs(SYNC_ROOT, exist_ok=True)
    sync_folder(token, drive_id, chg_folder["id"], SYNC_ROOT)

    # Count results
    total_files = sum(len(files) for _, _, files in os.walk(SYNC_ROOT))
    total_dirs = sum(len(dirs) for _, dirs, _ in os.walk(SYNC_ROOT))
    print(f"\n✅ Sync terminé!")
    print(f"📊 {total_files} fichiers, {total_dirs} dossiers téléchargés vers {SYNC_ROOT}")

if __name__ == "__main__":
    main()
