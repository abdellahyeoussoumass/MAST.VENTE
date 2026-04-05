import bcrypt
from database.db import session
from database.models import Utilisateur

def hasher_mot_de_passe(mot_de_passe):
    return bcrypt.hashpw(mot_de_passe.encode(), bcrypt.gensalt()).decode()

def verifier_mot_de_passe(mot_de_passe, hash_stocke):
    return bcrypt.checkpw(mot_de_passe.encode(), hash_stocke.encode())

def login(nom, mot_de_passe):
    user = session.query(Utilisateur).filter_by(nom=nom).first()
    if user and verifier_mot_de_passe(mot_de_passe, user.mot_de_passe):
        return user.role, user.nom  # ✅ retourne les deux
    return None, None