import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk
import os
from utils.theme import get_theme, basculer_mode
from utils.langue import get_text, changer_langue, get_langue
from database.db import session

APP_NAME  = "mast.vente"
NAVY      = "#1A2350"
NAVY2     = "#141C42"
ORANGE    = "#F5A623"
BLUE      = "#2979C8"

_BASE     = os.path.dirname(os.path.abspath(__file__))
_LOGO_DIR = os.path.join(_BASE, "..", "assets", "logos")


def _load_sidebar_logo(dark=True):
    """Charge le logo pour la sidebar (160×62 px)."""
    try:
        variant = "sidebar_dark" if dark else "sidebar_light"
        path    = os.path.join(_LOGO_DIR, f"logo_{variant}.png")
        img     = Image.open(path).resize((155, 60), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    except Exception as e:
        print(f"[main_window] Logo sidebar non chargé: {e}")
        return None


class MainWindow(ctk.CTk):
    def __init__(self, role, nom_utilisateur):
        super().__init__()
        self.role            = role
        self.nom_utilisateur = nom_utilisateur
        self.title(f"{APP_NAME} — {role.capitalize()}")
        self.geometry("1200x700")
        self.resizable(True, True)
        self._logo_photo = None   # référence gardée en mémoire

        # Backup automatique
        from utils.backup import backup_automatique, sauvegarder
        sauvegarder()
        backup_automatique(30)

        self.after(1000, self._verifier_stock)

        self._alpha = 0.0
        self.attributes("-alpha", 0.0)
        self.after(10, self._fade_in)
        self._construire()

    def _fade_in(self):
        if self._alpha < 1.0:
            self._alpha += 0.05
            self.attributes("-alpha", self._alpha)
            self.after(15, self._fade_in)

    def _construire(self):
        t        = get_theme()
        is_dark  = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")
        sidebar_bg = NAVY2 if is_dark else NAVY

        # ── SIDEBAR ───────────────────────────────────────
        self.sidebar = ctk.CTkFrame(
            self, width=220, corner_radius=0,
            fg_color=sidebar_bg)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # ── Logo dans la sidebar ──────────────────────────
        logo_frame = tk.Frame(self.sidebar, bg=sidebar_bg, height=80)
        logo_frame.pack(fill="x", pady=(10, 0))
        logo_frame.pack_propagate(False)

        photo = _load_sidebar_logo(dark=True)
        if photo:
            self._logo_photo = photo
            lbl = tk.Label(logo_frame, image=photo,
                           bg=sidebar_bg, bd=0)
            lbl.pack(expand=True)
        else:
            # Fallback texte
            tk.Label(logo_frame,
                     text="mast.vente",
                     bg=sidebar_bg, fg=ORANGE,
                     font=("Arial", 16, "bold")).pack(expand=True)

        # Ligne accent orange sous le logo
        tk.Frame(self.sidebar, bg=ORANGE, height=2).pack(
            fill="x", padx=10, pady=(4, 2))

        # Infos utilisateur
        ctk.CTkLabel(
            self.sidebar,
            text=f"👤  {self.nom_utilisateur.capitalize()}",
            font=("Arial", 11),
            text_color="#A0B4D0").pack(pady=(6, 0))

        ctk.CTkLabel(
            self.sidebar,
            text=f"🔑  {self.role.capitalize()}",
            font=("Arial", 10),
            text_color="#6A84A8").pack(pady=(2, 8))

        # Séparateur
        ctk.CTkFrame(self.sidebar, height=1,
                     fg_color="#2A3A6A").pack(fill="x", padx=12, pady=4)

        # ── Boutons menu ──────────────────────────────────
        boutons = []
        if self.role in ["directeur", "admin"]:
            boutons.append(("📊", get_text("dashboard"), "Dashboard"))

        boutons += [
            ("📦", get_text("produits"),      "Produits"),
            ("👥", get_text("clients"),        "Clients"),
            ("💰", get_text("ventes"),         "Ventes"),
            ("📋", get_text("devis"),          "Devis"),
            ("🛒", get_text("bons_commande"),  "Bons de Commande"),
            ("🧾", get_text("factures"),       "Factures"),
            ("🚚", get_text("livraisons"),     "Livraisons"),
            ("📖", get_text("guide"),          "Guide"),
            ("📅", get_text("historique"),     "Historique"),
            ("📊", "Rapport",                  "Rapport"),
        ]

        self.boutons_widgets = {}
        for emoji, label, nom in boutons:
            btn = ctk.CTkButton(
                self.sidebar,
                text=f"{emoji}  {label}",
                width=200, height=40,
                font=("Arial", 12),
                fg_color="transparent",
                hover_color="#1E3580",
                text_color="#B0C4DE",
                anchor="w",
                corner_radius=8,
                command=lambda n=nom: self.changer_section(n)
            )
            btn.pack(pady=1, padx=8)
            self.boutons_widgets[nom] = btn

        # ── Bas sidebar ───────────────────────────────────
        ctk.CTkFrame(self.sidebar, height=1,
                     fg_color="#2A3A6A").pack(
            fill="x", padx=12, pady=8, side="bottom")

        # Déconnexion
        ctk.CTkButton(
            self.sidebar,
            text="🚪  Déconnexion",
            width=200, height=40,
            font=("Arial", 12),
            fg_color="#C62828",
            hover_color="#8B0000",
            text_color=t["text_light"],
            corner_radius=8,
            command=self.deconnecter
        ).pack(side="bottom", pady=(0, 4), padx=8)

        # Sauvegarde
        ctk.CTkButton(
            self.sidebar,
            text="💾  Sauvegarder",
            width=200, height=38,
            font=("Arial", 11),
            fg_color="#00838F",
            hover_color="#006064",
            text_color=t["text_light"],
            corner_radius=8,
            command=self._backup_manuel
        ).pack(side="bottom", pady=(0, 2), padx=8)

        # Thème
        ctk.CTkButton(
            self.sidebar,
            text="🌙 / ☀️  Thème",
            width=200, height=38,
            font=("Arial", 11),
            fg_color="#2A3A6A",
            hover_color="#1E2D5A",
            text_color="#A0B4D0",
            corner_radius=8,
            command=self.changer_theme
        ).pack(side="bottom", pady=(0, 2), padx=8)

        # Sélecteur langue
        langue_frame = tk.Frame(self.sidebar, bg=sidebar_bg)
        langue_frame.pack(side="bottom", pady=4, padx=8, fill="x")

        for code, drapeau in [("fr", "🇫🇷"), ("en", "🇬🇧"), ("ar", "🇲🇦")]:
            ctk.CTkButton(
                langue_frame, text=drapeau,
                width=55, height=32,
                font=("Arial", 16),
                fg_color="#1E2D5A" if get_langue() == code else "transparent",
                hover_color="#2A3A6A",
                text_color=ORANGE if get_langue() == code else "#A0B4D0",
                corner_radius=6,
                command=lambda l=code: self._changer_langue(l)
            ).pack(side="left", padx=2)

        # ── ZONE CONTENU ──────────────────────────────────
        self.content = ctk.CTkFrame(self, fg_color=t["bg"])
        self.content.pack(side="right", fill="both", expand=True)

        self.after(1000, self._verifier_stock)
        self._afficher_bienvenue()

    # ──────────────────────────────────────────────────────
    def _changer_langue(self, langue):
        changer_langue(langue)
        self.destroy()
        MainWindow(self.role, self.nom_utilisateur).mainloop()

    def _backup_manuel(self):
        from utils.backup import sauvegarder
        from tkinter import messagebox
        sauvegarder()
        messagebox.showinfo(
            "✅ Sauvegarde",
            "Base de données sauvegardée avec succès !\n📁 Dossier : backups/")

    def _verifier_stock(self):
        from database.models import Produit
        from tkinter import messagebox
        produits_faibles = session.query(Produit).filter(
            Produit.quantite <= 5).all()
        if produits_faibles:
            msg = "⚠️ Stock faible pour :\n\n"
            for p in produits_faibles:
                msg += f"• {p.nom} — {p.quantite} unité(s)\n"
            messagebox.showwarning("🔔 Alerte Stock", msg)

    def _afficher_bienvenue(self):
        t = get_theme()
        for w in self.content.winfo_children():
            w.destroy()

        # Essayer d'afficher le logo sur l'écran de bienvenue
        is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")
        try:
            variant = "sidebar_dark" if is_dark else "sidebar_light"
            path    = os.path.join(_LOGO_DIR, f"logo_{variant}.png")
            img     = Image.open(path).resize((220, 85), Image.LANCZOS)
            photo   = ImageTk.PhotoImage(img)
            self._welcome_photo = photo
            lbl = tk.Label(self.content, image=photo, bg=t["bg"], bd=0)
            lbl.pack(pady=(120, 10))
        except Exception:
            ctk.CTkLabel(
                self.content,
                text="mast.vente",
                font=("Arial", 36, "bold"),
                text_color=ORANGE).pack(pady=(120, 10))

        ctk.CTkLabel(
            self.content,
            text=f"{get_text('bienvenue')}  {self.nom_utilisateur.capitalize()} !",
            font=("Arial", 22, "bold"),
            text_color=NAVY).pack(pady=(0, 6))

        ctk.CTkLabel(
            self.content,
            text=get_text("selection_section"),
            font=("Arial", 13),
            text_color="gray").pack(pady=4)

    def changer_section(self, nom):
        t = get_theme()
        # Reset couleurs boutons
        for n, btn in self.boutons_widgets.items():
            btn.configure(fg_color="transparent",
                          text_color="#B0C4DE")
        # Activer bouton sélectionné
        if nom in self.boutons_widgets:
            self.boutons_widgets[nom].configure(
                fg_color="#1E3580",
                text_color="#FFFFFF")

        for widget in self.content.winfo_children():
            widget.destroy()

        if nom == "Dashboard":
            from interfaces.dashboard import afficher_dashboard
            afficher_dashboard(self.content)
        elif nom == "Clients":
            from interfaces.clients import afficher_clients
            afficher_clients(self.content)
        elif nom == "Rapport":
            from interfaces.rapport import afficher_rapport
            afficher_rapport(self.content)
        elif nom == "Produits":
            from interfaces.produits import afficher_produits
            afficher_produits(self.content)
        elif nom == "Ventes":
            from interfaces.ventes import afficher_ventes
            afficher_ventes(self.content)
        elif nom == "Devis":
            from interfaces.devis import afficher_devis
            afficher_devis(self.content)
        elif nom == "Bons de Commande":
            from interfaces.bons_commande import afficher_bons_commande
            afficher_bons_commande(self.content)
        elif nom == "Factures":
            from interfaces.factures import afficher_factures
            afficher_factures(self.content)
        elif nom == "Livraisons":
            from interfaces.livraisons import afficher_livraisons
            afficher_livraisons(self.content)
        elif nom == "Guide":
            from interfaces.guide import afficher_guide
            afficher_guide(self.content)
        elif nom == "Historique":
            from interfaces.historique import afficher_historique
            afficher_historique(self.content)

    def changer_theme(self):
        basculer_mode()
        self.destroy()
        MainWindow(self.role, self.nom_utilisateur).mainloop()

    def deconnecter(self):
        self.destroy()
        from interfaces.login import LoginPage
        LoginPage().mainloop()
