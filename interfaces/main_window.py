import customtkinter as ctk
from utils.theme import get_theme, basculer_mode
from utils.langue import get_text, changer_langue, get_langue
from database.db import session

class MainWindow(ctk.CTk):
    def __init__(self, role, nom_utilisateur):
        super().__init__()
        self.role = role
        self.nom_utilisateur = nom_utilisateur
        self.title(f"VentePro — {role.capitalize()}")
        self.geometry("1200x700")
        self.resizable(True, True)

        # Backup automatique au démarrage
        from utils.backup import backup_automatique, sauvegarder
        sauvegarder()
        backup_automatique(30)

        # Vérifier stock faible
        self.after(1000, self._verifier_stock)

        # Animation ouverture
        self._alpha = 0.0
        self.attributes("-alpha", self._alpha)
        self.after(10, self._fade_in)

        self._construire()

    def _fade_in(self):
        if self._alpha < 1.0:
            self._alpha += 0.05
            self.attributes("-alpha", self._alpha)
            self.after(15, self._fade_in)

    def _construire(self):
        t = get_theme()

        # ---- Sidebar ----
        self.sidebar = ctk.CTkFrame(
            self, width=220, corner_radius=0, fg_color=t["sidebar"])
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Logo
        ctk.CTkLabel(
            self.sidebar,
            text="🧾 VentePro",
            font=("Arial", 20, "bold"),
            text_color="white"
        ).pack(pady=20)

        ctk.CTkLabel(
            self.sidebar,
            text=f"👤 {self.nom_utilisateur.capitalize()}",
            font=("Arial", 12),
            text_color="#FFE0B2"
        ).pack(pady=(0, 15))

        ctk.CTkLabel(
            self.sidebar,
            text=f"🔑 {self.role.capitalize()}",
            font=("Arial", 11),
            text_color="#FFE0B2"
        ).pack(pady=(0, 15))

        # Séparateur
        ctk.CTkFrame(self.sidebar, height=2,
                     fg_color="#FFE0B2").pack(fill="x", padx=15, pady=5)

        # Boutons menu
        boutons = []
        if self.role in ["directeur", "admin"]:
            boutons.append(("📊", get_text("dashboard"), "Dashboard"))

        boutons += [
            ("📦", get_text("produits"), "Produits"),
            ("👥", get_text("clients"), "Clients"),
            ("💰", get_text("ventes"), "Ventes"),
            ("📋", get_text("devis"), "Devis"),
            ("🛒", get_text("bons_commande"), "Bons de Commande"),
            ("🧾", get_text("factures"), "Factures"),
            ("🚚", get_text("livraisons"), "Livraisons"),
            ("📖", get_text("guide"), "Guide"),
            ("📅", get_text("historique"), "Historique"),
            ("📊", "Rapport", "Rapport"),
        ]

        self.boutons_widgets = {}
        for emoji, label, nom in boutons:
            btn = ctk.CTkButton(
                self.sidebar,
                text=f"{emoji}  {label}",
                width=200, height=42,
                font=("Arial", 13),
                fg_color="transparent",
                hover_color=t["sidebar_hover"],
                text_color="white",
                anchor="w",
                corner_radius=8,
                command=lambda n=nom: self.changer_section(n)
            )
            btn.pack(pady=2, padx=10)
            self.boutons_widgets[nom] = btn



        # Bouton dark/light
        ctk.CTkFrame(self.sidebar, height=2,
                     fg_color="#FFE0B2").pack(fill="x", padx=15, pady=10)

        ctk.CTkButton(
            self.sidebar,
            text="🌙 / ☀️  Thème",
            width=200, height=38,
            font=("Arial", 12),
            fg_color=t["secondary"],
            hover_color=t["primary_hover"],
            text_color="white",
            command=self.changer_theme
        ).pack(pady=3, padx=10)
        ctk.CTkButton(
            self.sidebar,
            text="💾  Sauvegarder",
            width=200, height=42,
            font=("Arial", 13),
            fg_color="#00838F",
            hover_color="#006064",
            text_color="white",
            command=self._backup_manuel
        ).pack(pady=3, padx=10)

        # Déconnexion
        ctk.CTkButton(
            self.sidebar,
            text="🚪  Déconnexion",
            width=200, height=42,
            font=("Arial", 13),
            fg_color=t["danger"],
            hover_color="#8B0000",
            text_color="white",
            command=self.deconnecter
        ).pack(side="bottom", pady=20, padx=10)

        # ---- Sélecteur de langue ----
        langue_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        langue_frame.pack(pady=5, padx=10, fill="x")

        for code, drapeau in [("fr", "🇫🇷"), ("en", "🇬🇧"), ("ar", "🇲🇦")]:
            ctk.CTkButton(
                langue_frame,
                text=drapeau,
                width=55, height=35,
                font=("Arial", 18),
                fg_color=t["primary"] if get_langue() == code else "transparent",
                hover_color=t["sidebar_hover"],
                command=lambda l=code: self._changer_langue(l)
            ).pack(side="left", padx=2)

        # ---- Zone contenu ----
        self.content = ctk.CTkFrame(self, fg_color=t["bg"])
        self.content.pack(side="right", fill="both", expand=True)
        # Vérifier stock faible au démarrage
        self.after(1000, self._verifier_stock)
        self._afficher_bienvenue()

    def _changer_langue(self, langue):
        changer_langue(langue)
        self.destroy()
        app = MainWindow(self.role, self.nom_utilisateur)
        app.mainloop()

    def _backup_manuel(self):
        from utils.backup import sauvegarder
        from tkinter import messagebox
        sauvegarder()
        messagebox.showinfo(
            "✅ Sauvegarde",
            "Base de données sauvegardée avec succès !\n📁 Dossier : backups/"
        )

    def _verifier_stock(self):
        from database.models import Produit
        seuil = 5
        produits_faibles = session.query(Produit).filter(
            Produit.quantite <= seuil
        ).all()

        if produits_faibles:
            message = "⚠️ Stock faible pour :\n\n"
            for p in produits_faibles:
                message += f"• {p.nom} — {p.quantite} unités restantes\n"

            from tkinter import messagebox
            messagebox.showwarning("🔔 Alerte Stock", message)

    def _afficher_bienvenue(self):
        t = get_theme()
        ctk.CTkLabel(
            self.content,
            text=get_text("bienvenue"),
            font=("Arial", 32, "bold"),
            text_color=t["primary"]
        ).pack(pady=(150, 10))

        ctk.CTkLabel(
            self.content,
            text=f"{self.nom_utilisateur.capitalize()} — {self.role.capitalize()}",
            font=("Arial", 16),
            text_color=t["secondary"]
        ).pack()

        ctk.CTkLabel(
            self.content,
            text=get_text("selection_section"),
            font=("Arial", 13),
            text_color="gray"
        ).pack(pady=10)

    def changer_section(self, nom):
        # Réinitialiser couleurs boutons
        t = get_theme()
        for n, btn in self.boutons_widgets.items():
            btn.configure(fg_color="transparent")

        # Highlight bouton actif
        if nom in self.boutons_widgets:
            self.boutons_widgets[nom].configure(fg_color=t["primary_hover"])

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

    def _afficher_profil(self):            # ← AJOUTE ICI
        for widget in self.content.winfo_children():
            widget.destroy()
        from interfaces.profil import afficher_profil
        afficher_profil(self.content, self.nom_utilisateur)

    def changer_theme(self):
        basculer_mode()
        self.destroy()
        app = MainWindow(self.role, self.nom_utilisateur)
        app.mainloop()

    def deconnecter(self):
        self.destroy()
        from interfaces.login import LoginPage
        app = LoginPage()
        app.mainloop()