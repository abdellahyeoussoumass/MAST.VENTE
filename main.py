import customtkinter as ctk
from database.db import session
from database.models import Utilisateur
from utils.auth import hasher_mot_de_passe

# Créer les utilisateurs si pas encore fait
if session.query(Utilisateur).count() == 0:
    utilisateurs = [
        Utilisateur(nom="directeur", mot_de_passe=hasher_mot_de_passe("1234"), role="directeur"),
        Utilisateur(nom="commercial", mot_de_passe=hasher_mot_de_passe("1234"), role="commercial"),
        Utilisateur(nom="admin", mot_de_passe=hasher_mot_de_passe("1234"), role="admin"),
    ]
    session.add_all(utilisateurs)
    session.commit()

# ---- Splash Screen puis Login ----
from interfaces.login import LoginPage

def lancer_app():
    app = LoginPage()
    app.mainloop()

# Créer fenêtre racine temporaire pour splash
root = ctk.CTk()
root.withdraw()  # Cacher la fenêtre racine

from interfaces.splash import SplashScreen

def apres_splash():
    root.destroy()  # Détruire fenêtre racine
    lancer_app()    # Lancer login

splash = SplashScreen(root)
# Quand splash se ferme → lancer login
root.after(3500, apres_splash)
root.mainloop()