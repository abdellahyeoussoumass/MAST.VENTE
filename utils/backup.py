import os
import shutil
from datetime import datetime
import threading

DOSSIER_BACKUP = "backups"

def creer_dossier_backup():
    if not os.path.exists(DOSSIER_BACKUP):
        os.makedirs(DOSSIER_BACKUP)

def sauvegarder():
    creer_dossier_backup()
    maintenant = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    nom_fichier = f"ventepro_backup_{maintenant}.db"
    destination = os.path.join(DOSSIER_BACKUP, nom_fichier)

    try:
        shutil.copy2("ventepro.db", destination)
        print(f"✅ Backup créé : {nom_fichier}")
        nettoyer_anciens_backups()
    except Exception as e:
        print(f"❌ Erreur backup : {e}")

def nettoyer_anciens_backups():
    # Garder seulement les 5 derniers backups
    fichiers = sorted([
        f for f in os.listdir(DOSSIER_BACKUP)
        if f.endswith(".db")
    ])
    while len(fichiers) > 5:
        os.remove(os.path.join(DOSSIER_BACKUP, fichiers[0]))
        fichiers.pop(0)
        print("🗑️ Ancien backup supprimé")

def backup_automatique(intervalle_minutes=30):
    def boucle():
        import time
        while True:
            time.sleep(intervalle_minutes * 60)
            sauvegarder()

    thread = threading.Thread(target=boucle, daemon=True)
    thread.start()
    print(f"⏰ Backup automatique activé toutes les {intervalle_minutes} minutes")