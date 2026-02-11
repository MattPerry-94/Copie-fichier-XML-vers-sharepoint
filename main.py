import sys
import os
from helpers import upload_driver_files_to_sharepoint

def main():
    try:
        success, _ = upload_driver_files_to_sharepoint()
        if not success:
            input("Appuyez sur Entrée pour fermer...")
            sys.exit(1)
        sys.exit(0)
    except Exception as e:
        print(f"Erreur: {e}")
        input("Appuyez sur Entrée pour fermer...")
        sys.exit(1)

if __name__ == "__main__":
    main()

