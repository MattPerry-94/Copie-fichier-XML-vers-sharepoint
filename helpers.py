import os
import sys
import time
import requests
import configparser
from msal import ConfidentialClientApplication
from urllib.parse import quote, unquote
from cryptography.fernet import Fernet

SHAREPOINT_FERNET_KEY = b'WIPVSOImL9zznQTiMR-PAC--e4m1-3F5EClMJ6oXbzg='
TENANT_ID = "2c2ab10a-7217-4b40-8641-3dade9eb9f84"

def resource_path(relative_path):
    """Retourne le chemin absolu d'un fichier, en tenant compte du mode exécutable ou développement."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(os.path.dirname(sys.executable), relative_path)
    else:
        return os.path.join(os.path.abspath("."), relative_path)

def lire_configuration(path="tools_XML.ini"):
    """Lit le fichier de configuration depuis le même dossier que l'exécutable."""
    config_path = resource_path(path)
    print(f"[LOG] Chemin du fichier de configuration : {config_path}")

    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Le fichier {path} est introuvable dans le dossier de l'application.")

    config_parser = configparser.ConfigParser()
    config_parser.read(config_path)

    fernet_sharepoint = Fernet(SHAREPOINT_FERNET_KEY)
    sharepoint_url = config_parser.get('SHAREPOINT', 'url', fallback='')

    source_folder_config = config_parser.get('DEFAULT', 'source_folder')
    if not source_folder_config:
        raise ValueError("Le chemin du dossier source n'est pas spécifié dans le fichier de configuration.")

    print(f"[LOG] Chemin du dossier source lu depuis le fichier de configuration : {source_folder_config}")

    config_default = {
        'source_folder': os.path.abspath(source_folder_config),
        'SHAREPOINT': {
            'url': sharepoint_url,
            'username': config_parser.get('SHAREPOINT', 'username', fallback=''),
            'library': config_parser.get('SHAREPOINT', 'library', fallback='LOG prog RGP'),
            'password': ''
        }
    }

    encrypted_sharepoint_secret = config_parser.get('SHAREPOINT', 'password', fallback='')
    if encrypted_sharepoint_secret:
        try:
            config_default['SHAREPOINT']['password'] = fernet_sharepoint.decrypt(
                encrypted_sharepoint_secret.encode()).decode()
        except Exception as e:
            raise ValueError(f"Erreur de déchiffrement du Client Secret SharePoint : {str(e)}")

    os.makedirs(config_default['source_folder'], exist_ok=True)
    return config_default

def upload_driver_files_to_sharepoint(ini_path="tools_XML.ini"):
    """Copie les fichiers vers SharePoint en respectant les critères définis."""
    config = lire_configuration(ini_path)
    stats = {'copies': 0, 'errors': 0, 'copied_files': []}
    source_folder = config['source_folder']
    client_id = config['SHAREPOINT']['username']
    client_secret = config['SHAREPOINT']['password']
    library_name = config['SHAREPOINT']['library']

    print(f"[LOG] Début de la copie des fichiers depuis {source_folder} vers la bibliothèque {library_name}")

    try:
        authority = f"https://login.microsoftonline.com/{TENANT_ID}"
        app = ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )

        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            print(f"[LOG] Échec du token: {result.get('error_description', 'Unknown error')}")
            return False, stats

        access_token = result["access_token"]
        headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/json'}

        # Accéder directement au site SharePoint
        site_response = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/ragnisas.sharepoint.com:/sites/Oprations",
            headers=headers
        )

        if site_response.status_code == 404:
            print("[LOG] Site Oprations non trouvé")
            return False, stats

        site_response.raise_for_status()
        site = site_response.json()
        site_id = site.get('id')
        if not site_id:
            print("[LOG] Impossible de récupérer l'ID du site SharePoint.")
            return False, stats

        print(f"[LOG] Site ID obtenu: {site_id}")

        # Récupération de la bibliothèque SharePoint
        drives_response = requests.get(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
            headers=headers
        )
        drives_response.raise_for_status()
        drives = drives_response.json().get('value', [])
        drive_id = next((d['id'] for d in drives if d['name'] == library_name), None)

        if not drive_id:
            print(f"[LOG] Bibliothèque '{library_name}' non trouvée.")
            print("[LOG] Bibliothèques disponibles :")
            for drive in drives:
                print(f"- {drive['name']} (ID: {drive['id']})")
            return False, stats

        print(f"[LOG] Bibliothèque '{library_name}' trouvée (ID: {drive_id})")

        # Liste des fichiers à copier
        all_files = os.listdir(source_folder)
        print(f"[LOG] Liste des fichiers dans le dossier source : {all_files}")

        valid_files = [
            f for f in all_files
            if f.endswith('.xml') or 'ORSTUP' in f.upper() or 'ORSTUL' in f.upper()
        ]

        if not valid_files:
            print("[LOG] Aucun fichier à copier trouvé dans le dossier source.")
            return False, stats

        # Upload des fichiers
        for fichier in valid_files:
            try:
                src = os.path.join(source_folder, fichier)
                if not os.path.isfile(src):
                    print(f"[LOG] {fichier} ignoré (n'est pas un fichier)")
                    continue

                # Vérifier si le fichier existe déjà sur SharePoint
                encoded_name = quote(fichier)
                check_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_name}"
                check_response = requests.get(check_url, headers=headers)

                if check_response.status_code == 200:
                    print(f"[LOG] Le fichier '{fichier}' existe déjà sur SharePoint → Ignoré.")
                    continue

                print(f"[LOG] Copie de {fichier} ({os.path.getsize(src)} octets)...")
                with open(src, 'rb') as f:
                    file_content = f.read()

                upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{encoded_name}:/content"
                upload_headers = {'Authorization': f'Bearer {access_token}', 'Content-Type': 'application/octet-stream'}

                response = requests.put(upload_url, headers=upload_headers, data=file_content)
                response.raise_for_status()
                print(f"✅ Copié: {fichier}")
                stats['copies'] += 1
                stats['copied_files'].append(fichier)
                time.sleep(1)

            except Exception as e:
                print(f"❌ Erreur pour {fichier}: {str(e)}")
                stats['errors'] += 1

        print(f"[LOG] Fin du traitement. {stats['copies']} fichiers copiés, {stats['errors']} erreurs.")
        return True, stats

    except Exception as e:
        print(f"[LOG] Erreur critique: {str(e)}")
        return False, stats
