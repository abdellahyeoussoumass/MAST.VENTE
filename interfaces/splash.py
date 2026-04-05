import customtkinter as ctk
from utils.theme import get_theme

class SplashScreen(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("")
        self.geometry("500x350")
        self.resizable(False, False)
        self.overrideredirect(True)  # Sans bordure

        # Centrer l'écran
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - 250
        y = (self.winfo_screenheight() // 2) - 175
        self.geometry(f"500x350+{x}+{y}")

        t = get_theme()
        self.configure(fg_color=t["primary"])

        # ---- Logo & Titre ----
        ctk.CTkLabel(
            self,
            text="🧾",
            font=("Arial", 60)
        ).pack(pady=(50, 5))

        ctk.CTkLabel(
            self,
            text="VentePro",
            font=("Arial", 40, "bold"),
            text_color="white"
        ).pack()

        ctk.CTkLabel(
            self,
            text="Sales Management v2.0",
            font=("Arial", 14),
            text_color="#FFE0B2"
        ).pack(pady=5)

        # ---- Barre de progression ----
        self.progress = ctk.CTkProgressBar(
            self, width=350, height=12,
            fg_color="#FFE0B2",
            progress_color="white"
        )
        self.progress.pack(pady=30)
        self.progress.set(0)

        self.label_status = ctk.CTkLabel(
            self,
            text="Initialisation...",
            font=("Arial", 12),
            text_color="#FFE0B2"
        )
        self.label_status.pack()

        # ---- Copyright ----
        ctk.CTkLabel(
            self,
            text="© 2025 VentePro — Tous droits réservés",
            font=("Arial", 10),
            text_color="#FFE0B2"
        ).pack(side="bottom", pady=15)

        # Démarrer animation
        self.etapes = [
            (0.15, "Chargement de la base de données..."),
            (0.35, "Initialisation des modules..."),
            (0.55, "Chargement de l'interface..."),
            (0.75, "Vérification des données..."),
            (0.90, "Préparation du tableau de bord..."),
            (1.00, "Bienvenue sur VentePro ! 🎉"),
        ]
        self.etape_actuelle = 0
        self.after(300, self._animer)

    def _animer(self):
        if self.etape_actuelle < len(self.etapes):
            valeur, message = self.etapes[self.etape_actuelle]
            self.progress.set(valeur)
            self.label_status.configure(text=message)
            self.etape_actuelle += 1
            self.after(500, self._animer)
        else:
            self.after(400, self.destroy)