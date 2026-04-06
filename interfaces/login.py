import customtkinter as ctk
from utils.auth import login
from utils.theme import get_theme, basculer_mode
import tkinter as tk
from PIL import Image, ImageTk
import os

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

NAVY    = "#1A2350"
NAVY2   = "#141C42"
ORANGE  = "#F5A623"
BLUE    = "#2979C8"
LBLUE   = "#5BB8F5"
WHITE   = "#FFFFFF"
CREAM   = "#F7F8FC"

_BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_LOGO_DIR = os.path.join(_BASE, "assets", "logos")

def _load_image(filename, size=None):
    try:
        path = os.path.join(_LOGO_DIR, filename)
        img  = Image.open(path)
        if size:
            img = img.resize(size, Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception as e:
        print(f"[login] Image '{filename}' non chargée: {e}")
        return None


class LoginPage(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("mast.vente")
        w, h = 980, 620
        self.geometry(f"{w}x{h}")
        self.resizable(False, False)
        self.configure(fg_color=CREAM)
        self.update_idletasks()
        x = (self.winfo_screenwidth()  // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self._alpha  = 0.0
        self._images = []
        self.attributes("-alpha", 0.0)
        self.after(10, self._fade_in)
        self._construire()

    def _fade_in(self):
        if self._alpha < 1.0:
            self._alpha = min(1.0, self._alpha + 0.07)
            self.attributes("-alpha", self._alpha)
            self.after(16, self._fade_in)

    def _construire(self):
        root_frame = tk.Frame(self, bg=CREAM)
        root_frame.pack(fill="both", expand=True)

        left = tk.Frame(root_frame, bg=NAVY, width=440)
        left.pack(side="left", fill="both")
        left.pack_propagate(False)
        self._build_left(left)

        right = tk.Frame(root_frame, bg=CREAM)
        right.pack(side="right", fill="both", expand=True)
        self._build_right(right)

    # ══════════════════════════════════════════════════════
    #   PANNEAU GAUCHE
    # ══════════════════════════════════════════════════════
    def _build_left(self, panel):
        cv = tk.Canvas(panel, bg=NAVY, highlightthickness=0,
                       width=440, height=620)
        cv.pack(fill="both", expand=True)

        # Formes décoratives
        cv.create_oval(-90, -90, 190, 190, fill=NAVY2, outline="")
        cv.create_oval(300, 480, 530, 710, fill=NAVY2, outline="")
        cv.create_oval(350, 55, 450, 155, fill="#1E3580", outline="")

        # Bandeau orange bas
        cv.create_rectangle(0, 588, 440, 620, fill=ORANGE, outline="")

        # ── LOGO IMAGE ────────────────────────────────────
        logo = _load_image("logo_login.png", size=(370, 144))
        if logo:
            self._images.append(logo)
            cv.create_image(220, 205, image=logo, anchor="center")
        else:
            # Fallback
            cv.create_text(220, 170, text="mast",
                           font=("Arial", 52, "bold"),
                           fill=WHITE, anchor="center")
            cv.create_text(220, 235, text=".vente",
                           font=("Arial", 40),
                           fill=ORANGE, anchor="center")

        # Tagline
        cv.create_text(220, 295,
                       text="Gestion Ventes & Suivi Commercial",
                       font=("Arial", 11, "italic"),
                       fill="#7A9CC0", anchor="center")

        # Séparateur orange
        cv.create_rectangle(70, 322, 370, 325,
                            fill=ORANGE, outline="")

        # Features
        features = [
            ("📦", "Produits & Stock"),
            ("💰", "Ventes & Facturation"),
            ("📋", "Devis & Bons de commande"),
            ("🚚", "Suivi des livraisons"),
            ("📊", "Dashboard & Rapports"),
        ]
        for i, (icon, label) in enumerate(features):
            y = 352 + i * 40
            # Fond pill
            cv.create_rectangle(52, y - 13, 396, y + 19,
                                 fill="#1E2D5A", outline="",
                                 width=0)
            # Bullet coloré
            color = ORANGE if i % 2 == 0 else LBLUE
            cv.create_oval(64, y - 5, 78, y + 9,
                           fill=color, outline="")
            # Icône + texte
            cv.create_text(92, y + 2, text=icon,
                           font=("Arial", 12), fill=WHITE, anchor="w")
            cv.create_text(116, y + 2, text=label,
                           font=("Arial", 10),
                           fill="#90AECB", anchor="w")

        # Copyright
        cv.create_text(220, 607,
                       text="© 2025 mast.vente",
                       font=("Arial", 8),
                       fill=NAVY2, anchor="center")

    # ══════════════════════════════════════════════════════
    #   PANNEAU DROIT
    # ══════════════════════════════════════════════════════
    def _build_right(self, panel):
        tk.Frame(panel, bg=CREAM, height=80).pack()

        # Titre
        header = tk.Frame(panel, bg=CREAM)
        header.pack(fill="x", padx=50)

        tk.Label(header, text="Connexion",
                 bg=CREAM, fg=NAVY,
                 font=("Arial", 30, "bold")).pack(anchor="w")
        tk.Label(header,
                 text="Bienvenue ! Accédez à votre espace de travail.",
                 bg=CREAM, fg="#7A8BAA",
                 font=("Arial", 11)).pack(anchor="w", pady=(5, 0))

        # Ligne accent
        tk.Frame(panel, bg=ORANGE, height=3).pack(
            fill="x", padx=50, pady=(16, 28))

        form = tk.Frame(panel, bg=CREAM)
        form.pack(fill="x", padx=50)

        # Nom utilisateur
        tk.Label(form, text="Nom d'utilisateur",
                 bg=CREAM, fg=NAVY,
                 font=("Arial", 11, "bold"),
                 anchor="w").pack(fill="x", pady=(0, 6))

        self.nom_entry = ctk.CTkEntry(
            form, height=48,
            placeholder_text="  Entrez votre identifiant",
            border_color=BLUE, border_width=2,
            corner_radius=10, fg_color=WHITE,
            text_color=NAVY, font=("Arial", 13))
        self.nom_entry.pack(fill="x", pady=(0, 18))

        # Mot de passe
        tk.Label(form, text="Mot de passe",
                 bg=CREAM, fg=NAVY,
                 font=("Arial", 11, "bold"),
                 anchor="w").pack(fill="x", pady=(0, 6))

        self.mdp_entry = ctk.CTkEntry(
            form, height=48,
            placeholder_text="  Entrez votre mot de passe",
            show="●", border_color=BLUE, border_width=2,
            corner_radius=10, fg_color=WHITE,
            text_color=NAVY, font=("Arial", 13))
        self.mdp_entry.pack(fill="x", pady=(0, 10))

        # Erreur
        self.erreur_label = tk.Label(
            form, text="", bg=CREAM,
            fg="#E53935", font=("Arial", 11), anchor="w")
        self.erreur_label.pack(fill="x", pady=(0, 14))

        # Bouton login
        self.btn_login = ctk.CTkButton(
            form, height=50,
            text="Se connecter  →",
            font=("Arial", 14, "bold"),
            fg_color=NAVY, hover_color=BLUE,
            text_color=WHITE, corner_radius=12,
            command=self.connecter)
        self.btn_login.pack(fill="x", pady=(0, 18))

        # Séparateur
        sep = tk.Frame(form, bg=CREAM)
        sep.pack(fill="x", pady=(0, 14))
        tk.Frame(sep, bg="#DDE3EE", height=1).pack(
            side="left", fill="x", expand=True, pady=9)
        tk.Label(sep, text="  ou  ", bg=CREAM, fg="#AAB4C8",
                 font=("Arial", 10)).pack(side="left")
        tk.Frame(sep, bg="#DDE3EE", height=1).pack(
            side="left", fill="x", expand=True, pady=9)

        # Bouton thème
        ctk.CTkButton(
            form, height=44,
            text="🌙  Mode Sombre  /  ☀️  Mode Clair",
            font=("Arial", 11), fg_color="transparent",
            hover_color="#E8EDF8", text_color=NAVY,
            border_color="#C8D0E0", border_width=2,
            corner_radius=10, command=self.changer_mode
        ).pack(fill="x")

        # Version
        tk.Label(panel, text="mast.vente v2.0",
                 bg=CREAM, fg="#BCC4D4",
                 font=("Arial", 9)).pack(side="bottom", pady=18)

        # Raccourcis
        self.mdp_entry.bind("<Return>", lambda e: self.connecter())
        self.nom_entry.bind("<Return>", lambda e: self.mdp_entry.focus())

    # ──────────────────────────────────────────────────────
    def changer_mode(self):
        basculer_mode()
        self.destroy()
        LoginPage().mainloop()

    def connecter(self):
        self.erreur_label.configure(text="")
        self.btn_login.configure(
            text="⏳  Connexion en cours...", state="disabled")
        self.after(450, self._verifier_login)

    def _verifier_login(self):
        nom = self.nom_entry.get().strip()
        mdp = self.mdp_entry.get()
        role, nom_utilisateur = login(nom, mdp)
        if role:
            self.destroy()
            from interfaces.main_window import MainWindow
            MainWindow(role, nom_utilisateur).mainloop()
        else:
            self.btn_login.configure(
                text="Se connecter  →", state="normal")
            self.erreur_label.configure(
                text="❌  Identifiant ou mot de passe incorrect")
            self._shake()

    def _shake(self):
        x, y  = self.winfo_x(), self.winfo_y()
        moves = [-12, 12, -9, 9, -6, 6, -3, 3, 0]
        for i, dx in enumerate(moves):
            self.after(i * 38, lambda d=dx: self.geometry(f"+{x + d}+{y}"))
