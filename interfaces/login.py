import customtkinter as ctk
from utils.auth import login
from utils.theme import get_theme, basculer_mode
import tkinter as tk

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class LoginPage(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("VentePro — Connexion")
        self.geometry("480x600")
        self.resizable(False, False)
        self.configure(fg_color=get_theme()["bg"])
        self._alpha = 0.0
        self.attributes("-alpha", self._alpha)
        self.after(10, self._fade_in)
        self._construire()

    def _fade_in(self):
        if self._alpha < 1.0:
            self._alpha += 0.05
            self.attributes("-alpha", self._alpha)
            self.after(20, self._fade_in)

    def _construire(self):
        t = get_theme()

        # ---- Header ----
        header = ctk.CTkFrame(self, fg_color=t["primary"], corner_radius=0, height=140)
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="🧾 VentePro",
            font=("Arial", 36, "bold"),
            text_color="white"
        ).pack(pady=(25, 0))

        ctk.CTkLabel(
            header,
            text="Sales Management v2.0",
            font=("Arial", 13),
            text_color="#FFE0B2"
        ).pack()

        # ---- Formulaire ----
        form = ctk.CTkFrame(self, fg_color=t["bg"], corner_radius=0)
        form.pack(fill="both", expand=True, padx=40)

        ctk.CTkLabel(form, text="Nom d'utilisateur",
                     font=("Arial", 13, "bold"),
                     text_color=t["primary"]).pack(pady=(30, 5), anchor="w")

        self.nom_entry = ctk.CTkEntry(
            form, width=380, height=45,
            placeholder_text="Entrez votre nom",
            border_color=t["primary"],
            font=("Arial", 13)
        )
        self.nom_entry.pack()

        ctk.CTkLabel(form, text="Mot de passe",
                     font=("Arial", 13, "bold"),
                     text_color=t["primary"]).pack(pady=(15, 5), anchor="w")

        self.mdp_entry = ctk.CTkEntry(
            form, width=380, height=45,
            placeholder_text="Entrez votre mot de passe",
            show="*", border_color=t["primary"],
            font=("Arial", 13)
        )
        self.mdp_entry.pack()

        self.erreur_label = ctk.CTkLabel(
            form, text="", text_color="red", font=("Arial", 12))
        self.erreur_label.pack(pady=8)

        self.btn_login = ctk.CTkButton(
            form,
            text="Se connecter",
            width=380, height=48,
            font=("Arial", 15, "bold"),
            fg_color=t["primary"],
            hover_color=t["primary_hover"],
            command=self.connecter
        )
        self.btn_login.pack(pady=5)

        # ---- Bouton dark/light ----
        ctk.CTkButton(
            form,
            text="🌙 Mode Sombre / ☀️ Mode Clair",
            width=380, height=38,
            font=("Arial", 12),
            fg_color=t["secondary"],
            hover_color=t["primary"],
            command=self.changer_mode
        ).pack(pady=10)

        # Entrée avec Enter
        self.mdp_entry.bind("<Return>", lambda e: self.connecter())

    def changer_mode(self):
        basculer_mode()
        self.destroy()
        app = LoginPage()
        app.mainloop()

    def connecter(self):
        # Animation bouton
        self.btn_login.configure(text="Connexion en cours...")
        self.after(500, self._verifier_login)

    def _verifier_login(self):
        nom = self.nom_entry.get()
        mdp = self.mdp_entry.get()
        role, nom_utilisateur = login(nom, mdp)

        if role:
            self.destroy()
            from interfaces.main_window import MainWindow
            app = MainWindow(role, nom_utilisateur)
            app.mainloop()
        else:
            self.btn_login.configure(text="Se connecter")
            self.erreur_label.configure(
                text="❌ Nom ou mot de passe incorrect !")
            self._shake()

    def _shake(self):
        x, y = self.winfo_x(), self.winfo_y()
        for i in range(6):
            self.after(i * 50, lambda dx=(-10 if i % 2 == 0 else 10):
                       self.geometry(f"+{x + dx}+{y}"))
        self.after(300, lambda: self.geometry(f"+{x}+{y}"))