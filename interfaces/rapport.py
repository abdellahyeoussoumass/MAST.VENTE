import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from utils.theme import get_theme
from datetime import date

def afficher_rapport(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t = get_theme()

    ctk.CTkLabel(
        parent,
        text="📊 Rapport Mensuel PDF",
        font=("Arial", 22, "bold"),
        text_color=t["primary"]
    ).pack(pady=20)

    # ---- Formulaire ----
    form = ctk.CTkFrame(parent, fg_color=t["card"],
                        corner_radius=15)
    form.pack(pady=20, padx=150, fill="x")

    ctk.CTkLabel(form, text="📅 Sélectionner la période",
                 font=("Arial", 16, "bold"),
                 text_color=t["primary"]).pack(pady=20)

    # Mois
    ctk.CTkLabel(form, text="Mois :",
                 font=("Arial", 13)).pack(pady=(5, 3))
    mois_noms = ["Janvier", "Février", "Mars", "Avril",
                 "Mai", "Juin", "Juillet", "Août",
                 "Septembre", "Octobre", "Novembre", "Décembre"]
    var_mois = tk.StringVar(value=mois_noms[date.today().month - 1])
    ctk.CTkComboBox(form, values=mois_noms,
                    variable=var_mois, width=250,
                    height=40).pack()

    # Année
    ctk.CTkLabel(form, text="Année :",
                 font=("Arial", 13)).pack(pady=(15, 3))
    annees = [str(a) for a in range(2021, date.today().year + 1)]
    var_annee = tk.StringVar(value=str(date.today().year))
    ctk.CTkComboBox(form, values=annees,
                    variable=var_annee, width=250,
                    height=40).pack()

    label_status = ctk.CTkLabel(form, text="",
                                font=("Arial", 12))
    label_status.pack(pady=10)

    def generer():
        try:
            mois = mois_noms.index(var_mois.get()) + 1
            annee = int(var_annee.get())

            label_status.configure(
                text="⏳ Génération en cours...",
                text_color="orange"
            )
            form.update()

            from utils.rapport import generer_rapport
            fichier = generer_rapport(mois, annee)

            label_status.configure(
                text=f"✅ Rapport généré : {fichier}",
                text_color="green"
            )
            messagebox.showinfo(
                "✅ Succès",
                f"Rapport généré avec succès !\n📄 Fichier : {fichier}"
            )
        except Exception as e:
            label_status.configure(
                text=f"❌ Erreur : {e}",
                text_color="red"
            )

    ctk.CTkButton(
        form,
        text="📊 Générer le Rapport PDF",
        height=48, width=250,
        font=("Arial", 14, "bold"),
        fg_color=t["primary"],
        hover_color=t["primary_hover"],
        command=generer
    ).pack(pady=20)