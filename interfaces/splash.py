import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk
import os

NAVY   = "#1A2350"
NAVY2  = "#141C42"
ORANGE = "#F5A623"
BLUE   = "#2979C8"
WHITE  = "#FFFFFF"

_BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_LOGO_DIR = os.path.join(_BASE, "assets", "logos")


class SplashScreen(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("")
        self.geometry("540x370")
        self.resizable(False, False)
        self.overrideredirect(True)

        self.update_idletasks()
        x = (self.winfo_screenwidth()  // 2) - 270
        y = (self.winfo_screenheight() // 2) - 185
        self.geometry(f"540x370+{x}+{y}")

        self.configure(fg_color=NAVY)
        self._images = []
        self._construire()

    def _construire(self):
        cv = tk.Canvas(self, bg=NAVY, highlightthickness=0,
                       width=540, height=370)
        cv.pack(fill="both", expand=True)

        # Formes décoratives
        cv.create_oval(-80, -80, 180, 180, fill=NAVY2, outline="")
        cv.create_oval(400, -40, 600, 160, fill="#1E3580", outline="")
        cv.create_oval(380, 300, 580, 450, fill=NAVY2, outline="")
        # Bandeau orange bas
        cv.create_rectangle(0, 346, 540, 370, fill=ORANGE, outline="")

        # ── LOGO ─────────────────────────────────────────
        try:
            path  = os.path.join(_LOGO_DIR, "logo_login.png")
            img   = Image.open(path).resize((360, 140), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self._images.append(photo)
            cv.create_image(270, 145, image=photo, anchor="center")
        except Exception:
            cv.create_text(270, 120, text="mast",
                           font=("Arial", 48, "bold"),
                           fill=WHITE, anchor="center")
            cv.create_text(270, 175, text=".vente",
                           font=("Arial", 36),
                           fill=ORANGE, anchor="center")

        # ── Barre de progression ──────────────────────────
        # Fond barre
        cv.create_rectangle(80, 236, 460, 252,
                            fill="#0E1730", outline="", width=0)
        cv.create_rectangle(82, 238, 458, 250,
                            fill="#1E2D5A", outline="", width=0)
        # Barre elle-même (sera animée)
        self._bar_item = cv.create_rectangle(
            82, 238, 82, 250, fill=ORANGE, outline="", width=0)
        self._bar_end  = cv.create_rectangle(
            82, 238, 82, 250, fill=BLUE, outline="", width=0)

        self._cv = cv

        # ── Label statut ─────────────────────────────────
        self._status_var = tk.StringVar(value="Initialisation...")
        status_lbl = tk.Label(self, textvariable=self._status_var,
                              bg=NAVY, fg="#7A9CC0",
                              font=("Arial", 11))
        status_lbl.place(x=270, y=262, anchor="center")
        self._status_lbl = status_lbl

        # Copyright
        cv.create_text(270, 356,
                       text="© 2025 mast.vente",
                       font=("Arial", 8), fill=NAVY2, anchor="center")

        # Démarrer animation
        self._etapes = [
            (0.15, "Chargement de la base de données..."),
            (0.30, "Initialisation des modules..."),
            (0.50, "Chargement de l'interface..."),
            (0.68, "Vérification des données..."),
            (0.85, "Préparation du tableau de bord..."),
            (1.00, "Bienvenue sur mast.vente ! 🎉"),
        ]
        self._etape_idx = 0
        self._progress  = 0.0
        self.after(250, self._animer)

    def _animer(self):
        if self._etape_idx < len(self._etapes):
            cible, message = self._etapes[self._etape_idx]
            self._progress  = cible
            self._status_var.set(message)
            self._etape_idx += 1
            # Mettre à jour la barre
            x_end = 82 + int((458 - 82) * cible)
            self._cv.coords(self._bar_item, 82, 238, x_end, 250)
            self.after(480, self._animer)
        else:
            self.after(350, self.destroy)
