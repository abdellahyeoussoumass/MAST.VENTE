import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from database.db import session
from database.models import Vente, Client, Produit
from datetime import date
from utils.theme import get_theme
import os

# ══════════════════════════════════════════════════════════
#   IMPORTS OPTIONNELS
# ══════════════════════════════════════════════════════════
try:
    import openpyxl
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    import pandas as pd
    PANDAS_OK = True
except ImportError:
    PANDAS_OK = False

import csv
import json


# ══════════════════════════════════════════════════════════
#   CONSTANTES
# ══════════════════════════════════════════════════════════
EN_TETES = ["ID", "Client", "Produit", "Date", "Quantite",
            "Prix", "Reduction (%)", "Montant Net", "Ville"]
CHAMPS   = ["id", "client", "produit", "date_vente", "quantite",
            "prix", "reduction", "montant_net", "ville"]


# ══════════════════════════════════════════════════════════
#   STYLE TREEVIEW
# ══════════════════════════════════════════════════════════
def _style_treeview(is_dark):
    style = ttk.Style()
    style.theme_use("clam")

    bg_tree = "#12122A" if is_dark else "#FFFFFF"
    bg_head = "#E65100"
    bg_sel  = "#FF6D00" if is_dark else "#FFB74D"
    fg_text = "#E8E0FF" if is_dark else "#1A1A1A"
    bg_row1 = "#1A1A35" if is_dark else "#FFFFFF"
    bg_row2 = "#12122A" if is_dark else "#FFF8F0"

    style.configure("VN.Treeview",
                    background=bg_tree, foreground=fg_text,
                    rowheight=32, fieldbackground=bg_tree,
                    borderwidth=0, font=("Arial", 10))
    style.configure("VN.Treeview.Heading",
                    background=bg_head, foreground="white",
                    font=("Arial", 10, "bold"),
                    relief="flat", borderwidth=0)
    style.map("VN.Treeview",
              background=[("selected", bg_sel)],
              foreground=[("selected", "#FFFFFF")])
    style.map("VN.Treeview.Heading",
              background=[("active", "#BF360C")])
    return bg_row1, bg_row2


# ══════════════════════════════════════════════════════════
#   HELPERS UI
# ══════════════════════════════════════════════════════════
def _fenetre_modale(titre, largeur=480, hauteur=600):
    t   = get_theme()
    win = tk.Toplevel()
    win.title(titre)
    win.geometry(f"{largeur}x{hauteur}")
    win.configure(bg=t["bg"])
    win.resizable(False, False)
    win.grab_set()

    hdr = tk.Frame(win, bg="#E65100", height=50)
    hdr.pack(fill="x")
    tk.Label(hdr, text=titre, bg="#E65100", fg="white",
             font=("Arial", 14, "bold")).pack(pady=12)

    body = tk.Frame(win, bg=t["bg"])
    body.pack(fill="both", expand=True, padx=24, pady=16)
    return win, body, t


def _lbl(body, text, t):
    tk.Label(body, text=text, bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(
             fill="x", pady=(10, 2))


def _entry(body, t, default=""):
    e = tk.Entry(body, font=("Arial", 11),
                 bg=t["card"], fg=t["text"],
                 insertbackground=t["text"],
                 relief="flat", bd=6)
    e.pack(fill="x", ipady=4)
    if default != "":
        e.insert(0, str(default))
    return e


def _combo(body, t, values, default=""):
    var = tk.StringVar(value=default)
    cb  = ttk.Combobox(body, textvariable=var, values=values,
                       font=("Arial", 11), state="readonly")
    cb.pack(fill="x", ipady=4)
    cb.var = var
    return cb


def _btn_ok(parent, text, cmd, color="#E65100", hover="#BF360C"):
    return tk.Button(
        parent, text=text, command=cmd,
        bg=color, fg="white",
        font=("Arial", 12, "bold"),
        relief="flat", cursor="hand2",
        padx=16, pady=9,
        activebackground=hover,
        activeforeground="white",
    )


def _sep(body, t):
    tk.Frame(body,
             bg="#2A2A4A" if t["bg"] in ("#1A1A2E", "#0F0F1A") else "#E8D8C8",
             height=1).pack(fill="x", pady=(10, 0))


# ══════════════════════════════════════════════════════════
#   TRI COLONNES
# ══════════════════════════════════════════════════════════
_tri_etat = {}

def _trier(table, col):
    rows = [(table.set(r, col), r) for r in table.get_children("")]
    rev  = _tri_etat.get(col, False)
    try:
        rows.sort(key=lambda x: float(
            x[0].replace(",", "").replace("%", "")), reverse=rev)
    except ValueError:
        rows.sort(key=lambda x: x[0].lower(), reverse=rev)
    for i, (_, r) in enumerate(rows):
        table.move(r, "", i)
    _tri_etat[col] = not rev


# ══════════════════════════════════════════════════════════
#   MENU DEROULANT IMPORTER
# ══════════════════════════════════════════════════════════
def _show_dropdown_import(btn_widget, table, t):
    is_dark  = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")
    bg_menu  = "#1A1A35" if is_dark else "#FFFFFF"
    bg_hover = "#252545" if is_dark else "#FFF0E5"
    border_c = "#2A2A4A" if is_dark else "#E8D8C8"

    items = [
        ("📊  Excel (.xlsx)", "#1565C0", lambda: _importer(table, "excel")),
        ("📄  CSV (.csv)",    "#00838F", lambda: _importer(table, "csv")),
        ("📋  JSON (.json)",  "#6A1B9A", lambda: _importer(table, "json")),
    ]

    menu = tk.Toplevel()
    menu.overrideredirect(True)
    menu.configure(bg=border_c)
    menu.attributes("-topmost", True)

    btn_widget.update_idletasks()
    x = btn_widget.winfo_rootx()
    y = btn_widget.winfo_rooty() + btn_widget.winfo_height() + 2
    menu.geometry(f"210x{len(items) * 40 + 16}+{x}+{y}")

    inner = tk.Frame(menu, bg=bg_menu, bd=0)
    inner.pack(fill="both", expand=True, padx=1, pady=1)

    tk.Label(inner, text="Choisir le format :",
             bg=bg_menu, fg="#FF8F00",
             font=("Arial", 8, "bold"),
             anchor="w").pack(fill="x", padx=10, pady=(6, 2))

    def _fermer():
        try:
            menu.destroy()
        except Exception:
            pass

    for label, couleur, cmd_fn in items:
        f   = tk.Frame(inner, bg=bg_menu, cursor="hand2")
        f.pack(fill="x", padx=4, pady=1)
        lbl = tk.Label(f, text=label, bg=bg_menu, fg=couleur,
                       font=("Arial", 10, "bold"),
                       anchor="w", padx=12, pady=8)
        lbl.pack(fill="x")

        def _enter(e, fr=f, lb=lbl):
            fr.configure(bg=bg_hover); lb.configure(bg=bg_hover)
        def _leave(e, fr=f, lb=lbl):
            fr.configure(bg=bg_menu);  lb.configure(bg=bg_menu)
        def _click(e, fn=cmd_fn):
            _fermer(); fn()

        for w in (f, lbl):
            w.bind("<Enter>",    _enter)
            w.bind("<Leave>",    _leave)
            w.bind("<Button-1>", _click)

    menu.bind("<FocusOut>", lambda e: _fermer())
    menu.focus_set()


# ══════════════════════════════════════════════════════════
#   AFFICHER VENTES
# ══════════════════════════════════════════════════════════
def afficher_ventes(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t       = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")

    # En-tete
    header = ctk.CTkFrame(parent, fg_color=t["card"], corner_radius=0)
    header.pack(fill="x")
    ctk.CTkLabel(header, text="💰  Gestion des Ventes",
                 font=("Arial", 22, "bold"),
                 text_color="#E65100").pack(side="left", padx=24, pady=14)

    # Barre outils
    toolbar = ctk.CTkFrame(parent, fg_color="transparent")
    toolbar.pack(fill="x", padx=20, pady=(10, 4))

    btns = [
        ("➕  Ajouter",    "#E65100", "#BF360C", lambda: ajouter_vente(table)),
        ("✏️  Modifier",   "#FF8F00", "#E65100", lambda: modifier_vente(table)),
        ("🗑️  Supprimer",  "#C62828", "#8B0000", lambda: supprimer_vente(table)),
    ]
    for txt, fg, hover, cmd in btns:
        ctk.CTkButton(toolbar, text=txt, width=128, height=36,
                      fg_color=fg, hover_color=hover,
                      font=("Arial", 11, "bold"),
                      corner_radius=8,
                      command=cmd).pack(side="left", padx=4)

    # Bouton Importer dropdown
    btn_imp = ctk.CTkButton(
        toolbar, text="📥  Importer ▾", width=138, height=36,
        fg_color="#1565C0", hover_color="#0D47A1",
        font=("Arial", 11, "bold"), corner_radius=8,
        command=lambda: _show_dropdown_import(btn_imp, table, t)
    )
    btn_imp.pack(side="left", padx=4)

    # Compteur + Recherche
    count_var = tk.StringVar()
    ctk.CTkLabel(toolbar, textvariable=count_var,
                 font=("Arial", 11),
                 text_color="#FF8F00").pack(side="right", padx=10)

    recherche_var = tk.StringVar()
    ctk.CTkEntry(toolbar,
                 placeholder_text="🔍  Rechercher...",
                 width=240, height=36,
                 textvariable=recherche_var,
                 font=("Arial", 11)).pack(side="right", padx=5)
    recherche_var.trace("w",
        lambda *a: rechercher(recherche_var.get(), table))

    # Tableau
    frame_table = ctk.CTkFrame(parent, corner_radius=12)
    frame_table.pack(fill="both", expand=True, padx=20, pady=(4, 16))

    bg_row1, bg_row2 = _style_treeview(is_dark)

    colonnes = ("ID", "Client", "Produit", "Date",
                "Qte", "Prix", "Reduction", "Montant Net", "Ville")

    table = ttk.Treeview(frame_table, columns=colonnes,
                         show="headings", height=22,
                         style="VN.Treeview",
                         selectmode="browse")

    largeurs = {"ID": 45, "Client": 150, "Produit": 150,
                "Date": 100, "Qte": 55, "Prix": 90,
                "Reduction": 85, "Montant Net": 110, "Ville": 110}

    for col in colonnes:
        table.heading(col, text=col,
                      command=lambda c=col: _trier(table, c))
        table.column(col, width=largeurs.get(col, 100), anchor="center")

    table.tag_configure("pair",   background=bg_row1)
    table.tag_configure("impair", background=bg_row2)
    # Couleur montant net selon valeur
    table.tag_configure("net_haut",  foreground="#43A047")
    table.tag_configure("net_moyen", foreground="#FF8F00")

    sb_v = ttk.Scrollbar(frame_table, orient="vertical",   command=table.yview)
    sb_h = ttk.Scrollbar(frame_table, orient="horizontal", command=table.xview)
    table.configure(yscroll=sb_v.set, xscroll=sb_h.set)
    sb_v.pack(side="right",  fill="y")
    sb_h.pack(side="bottom", fill="x")
    table.pack(fill="both", expand=True, padx=2, pady=2)

    table.bind("<Double-1>", lambda e: modifier_vente(table))

    charger_ventes(table, count_var)
    return table


# ══════════════════════════════════════════════════════════
#   CHARGER VENTES
# ══════════════════════════════════════════════════════════
def charger_ventes(table, count_var=None):
    for row in table.get_children():
        table.delete(row)

    ventes = session.query(Vente).order_by(Vente.id.desc()).all()

    # Calculer montant max pour colorisation relative
    montants = [float(v.montant_net or 0) for v in ventes]
    max_net  = max(montants) if montants else 1

    for i, v in enumerate(ventes):
        client  = session.query(Client).filter_by(id=v.client_id).first()
        produit = session.query(Produit).filter_by(id=v.produit_id).first()

        net    = float(v.montant_net or 0)
        tags   = ["pair" if i % 2 == 0 else "impair"]
        if max_net > 0:
            if net >= max_net * 0.6:
                tags.append("net_haut")
            elif net >= max_net * 0.3:
                tags.append("net_moyen")

        table.insert("", "end", tags=tuple(tags), values=(
            v.id,
            client.nom  if client  else "N/A",
            produit.nom if produit else "N/A",
            str(v.date_vente),
            v.quantite,
            f"{float(v.prix):,.2f}"         if v.prix        else "0.00",
            f"{float(v.reduction):.1f}%"    if v.reduction   else "0%",
            f"{net:,.2f}",
            v.ville or "-",
        ))

    if count_var is not None:
        total = sum(montants)
        count_var.set(f"{len(ventes)} vente(s)  |  Total : {total:,.2f} MAD")


# ══════════════════════════════════════════════════════════
#   RECHERCHER
# ══════════════════════════════════════════════════════════
def rechercher(texte, table):
    for row in table.get_children():
        table.delete(row)

    texte_low = texte.lower()
    ventes    = session.query(Vente).all()

    for i, v in enumerate(ventes):
        client  = session.query(Client).filter_by(id=v.client_id).first()
        produit = session.query(Produit).filter_by(id=v.produit_id).first()
        nom_c   = client.nom  if client  else ""
        nom_p   = produit.nom if produit else ""

        if (texte_low in nom_c.lower()
                or texte_low in nom_p.lower()
                or texte_low in (v.ville or "").lower()
                or texte_low in str(v.date_vente).lower()):

            tags = ("pair" if i % 2 == 0 else "impair",)
            table.insert("", "end", tags=tags, values=(
                v.id,
                nom_c or "N/A",
                nom_p or "N/A",
                str(v.date_vente),
                v.quantite,
                f"{float(v.prix):,.2f}"      if v.prix      else "0.00",
                f"{float(v.reduction):.1f}%" if v.reduction else "0%",
                f"{float(v.montant_net):,.2f}" if v.montant_net else "0.00",
                v.ville or "-",
            ))


# ══════════════════════════════════════════════════════════
#   AJOUTER VENTE
# ══════════════════════════════════════════════════════════
def ajouter_vente(table):
    clients  = session.query(Client).all()
    produits = session.query(Produit).all()

    win, body, t = _fenetre_modale("➕  Nouvelle Vente", 480, 620)

    _lbl(body, "Client", t)
    cb_client = _combo(body, t, [c.nom for c in clients])

    _lbl(body, "Produit", t)
    cb_produit = _combo(body, t, [p.nom for p in produits])

    _lbl(body, "Ville", t)
    e_ville = _entry(body, t)

    # Quantite + Reduction sur la meme ligne
    row2 = tk.Frame(body, bg=t["bg"])
    row2.pack(fill="x", pady=(10, 0))

    col1 = tk.Frame(row2, bg=t["bg"])
    col1.pack(side="left", expand=True, fill="x", padx=(0, 8))
    tk.Label(col1, text="Quantite", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_qte = tk.Entry(col1, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6)
    e_qte.pack(fill="x", ipady=4)

    col2 = tk.Frame(row2, bg=t["bg"])
    col2.pack(side="left", expand=True, fill="x")
    tk.Label(col2, text="Reduction (%)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_red = tk.Entry(col2, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6)
    e_red.insert(0, "0")
    e_red.pack(fill="x", ipady=4)

    # Apercu calcul
    lbl_preview = tk.Label(body, text="",
                           bg=t["bg"], fg="#FF8F00",
                           font=("Arial", 11, "bold"), anchor="w")
    lbl_preview.pack(fill="x", pady=(8, 0))

    def maj_preview(*_):
        try:
            produit = session.query(Produit).filter_by(
                nom=cb_produit.var.get()).first()
            if produit and e_qte.get():
                qte  = int(e_qte.get())
                prix = float(produit.prix_ht)
                red  = float(e_red.get() or 0)
                brut = prix * qte
                red_mnt = brut * red / 100
                net  = round(brut - red_mnt, 2)
                lbl_preview.config(
                    text=f"Prix HT : {prix:,.2f}  ×  {qte}  −  {red_mnt:,.2f} MAD  =  Net : {net:,.2f} MAD")
        except Exception:
            pass

    cb_produit.var.trace("w", maj_preview)
    e_qte.bind("<KeyRelease>", maj_preview)
    e_red.bind("<KeyRelease>", maj_preview)

    _sep(body, t)
    btn_frame = tk.Frame(body, bg=t["bg"])
    btn_frame.pack(fill="x", pady=(12, 4))

    def sauvegarder():
        try:
            client  = session.query(Client).filter_by(
                nom=cb_client.var.get()).first()
            produit = session.query(Produit).filter_by(
                nom=cb_produit.var.get()).first()
            if not client:
                messagebox.showerror("Erreur", "Selectionnez un client.", parent=win)
                return
            if not produit:
                messagebox.showerror("Erreur", "Selectionnez un produit.", parent=win)
                return

            qte  = int(e_qte.get())
            prix = float(produit.prix_ht)
            red  = float(e_red.get() or 0)
            net  = round(prix * qte * (1 - red / 100), 2)

            v = Vente(
                client_id=client.id,
                produit_id=produit.id,
                date_vente=date.today(),
                quantite=qte,
                prix=prix,
                reduction=red,
                montant_net=net,
                ville=e_ville.get().strip(),
            )
            session.add(v)
            session.commit()
            charger_ventes(table)
            win.destroy()
            messagebox.showinfo("Succes", "Vente ajoutee avec succes !")
        except Exception as ex:
            messagebox.showerror("Erreur", f"Erreur : {ex}", parent=win)

    _btn_ok(btn_frame, "💾  Sauvegarder", sauvegarder).pack(
        side="left", expand=True, fill="x", padx=(0, 6))
    tk.Button(btn_frame, text="Annuler", command=win.destroy,
              bg=t["card"], fg=t["text"], font=("Arial", 11),
              relief="flat", cursor="hand2", pady=9).pack(
              side="left", expand=True, fill="x")


# ══════════════════════════════════════════════════════════
#   MODIFIER VENTE
# ══════════════════════════════════════════════════════════
def modifier_vente(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez une vente a modifier !")
        return

    valeurs = table.item(sel[0])["values"]
    vente   = session.query(Vente).filter_by(id=valeurs[0]).first()
    if not vente:
        return

    clients  = session.query(Client).all()
    produits = session.query(Produit).all()
    client_actuel  = session.query(Client).filter_by(id=vente.client_id).first()
    produit_actuel = session.query(Produit).filter_by(id=vente.produit_id).first()

    win, body, t = _fenetre_modale("✏️  Modifier la Vente", 480, 640)

    _lbl(body, "Client", t)
    cb_client = _combo(body, t, [c.nom for c in clients],
                       default=client_actuel.nom if client_actuel else "")

    _lbl(body, "Produit", t)
    cb_produit = _combo(body, t, [p.nom for p in produits],
                        default=produit_actuel.nom if produit_actuel else "")

    _lbl(body, "Ville", t)
    e_ville = _entry(body, t, default=vente.ville or "")

    # Quantite + Reduction cote a cote
    row2 = tk.Frame(body, bg=t["bg"])
    row2.pack(fill="x", pady=(10, 0))

    col1 = tk.Frame(row2, bg=t["bg"])
    col1.pack(side="left", expand=True, fill="x", padx=(0, 8))
    tk.Label(col1, text="Quantite", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_qte = tk.Entry(col1, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6)
    e_qte.insert(0, str(vente.quantite or 1))
    e_qte.pack(fill="x", ipady=4)

    col2 = tk.Frame(row2, bg=t["bg"])
    col2.pack(side="left", expand=True, fill="x")
    tk.Label(col2, text="Reduction (%)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_red = tk.Entry(col2, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6)
    e_red.insert(0, str(float(vente.reduction or 0)))
    e_red.pack(fill="x", ipady=4)

    # Apercu calcul
    lbl_preview = tk.Label(body, text="",
                           bg=t["bg"], fg="#FF8F00",
                           font=("Arial", 11, "bold"), anchor="w")
    lbl_preview.pack(fill="x", pady=(8, 0))

    def maj_preview(*_):
        try:
            produit = session.query(Produit).filter_by(
                nom=cb_produit.var.get()).first()
            if produit and e_qte.get():
                qte  = int(e_qte.get())
                prix = float(produit.prix_ht)
                red  = float(e_red.get() or 0)
                brut = prix * qte
                red_mnt = brut * red / 100
                net  = round(brut - red_mnt, 2)
                lbl_preview.config(
                    text=f"Prix HT : {prix:,.2f}  ×  {qte}  −  {red_mnt:,.2f}  =  Net : {net:,.2f} MAD")
        except Exception:
            pass

    cb_produit.var.trace("w", maj_preview)
    e_qte.bind("<KeyRelease>", maj_preview)
    e_red.bind("<KeyRelease>", maj_preview)
    maj_preview()

    # Date
    _lbl(body, "Date de vente", t)
    e_date = _entry(body, t, default=str(vente.date_vente or date.today()))

    _sep(body, t)
    btn_frame = tk.Frame(body, bg=t["bg"])
    btn_frame.pack(fill="x", pady=(12, 4))

    def sauvegarder():
        try:
            client  = session.query(Client).filter_by(
                nom=cb_client.var.get()).first()
            produit = session.query(Produit).filter_by(
                nom=cb_produit.var.get()).first()
            if not client:
                messagebox.showerror("Erreur", "Selectionnez un client.", parent=win)
                return
            if not produit:
                messagebox.showerror("Erreur", "Selectionnez un produit.", parent=win)
                return

            qte  = int(e_qte.get())
            prix = float(produit.prix_ht)
            red  = float(e_red.get() or 0)
            net  = round(prix * qte * (1 - red / 100), 2)

            vente.client_id  = client.id
            vente.produit_id = produit.id
            vente.ville      = e_ville.get().strip()
            vente.quantite   = qte
            vente.prix       = prix
            vente.reduction  = red
            vente.montant_net= net
            try:
                from datetime import datetime
                vente.date_vente = datetime.strptime(
                    e_date.get().strip(), "%Y-%m-%d").date()
            except Exception:
                pass

            session.commit()
            charger_ventes(table)
            win.destroy()
            messagebox.showinfo("Succes", "Vente modifiee avec succes !")
        except Exception as ex:
            messagebox.showerror("Erreur", f"Erreur : {ex}", parent=win)

    _btn_ok(btn_frame, "💾  Sauvegarder", sauvegarder).pack(
        side="left", expand=True, fill="x", padx=(0, 6))
    tk.Button(btn_frame, text="Annuler", command=win.destroy,
              bg=t["card"], fg=t["text"], font=("Arial", 11),
              relief="flat", cursor="hand2", pady=9).pack(
              side="left", expand=True, fill="x")


# ══════════════════════════════════════════════════════════
#   SUPPRIMER VENTE
# ══════════════════════════════════════════════════════════
def supprimer_vente(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez une vente a supprimer !")
        return

    valeurs = table.item(sel[0])["values"]
    vente   = session.query(Vente).filter_by(id=valeurs[0]).first()
    if not vente:
        return

    client  = session.query(Client).filter_by(id=vente.client_id).first()
    produit = session.query(Produit).filter_by(id=vente.produit_id).first()

    if messagebox.askyesno(
        "Confirmation",
        f"Supprimer cette vente ?\n"
        f"Client  : {client.nom  if client  else 'N/A'}\n"
        f"Produit : {produit.nom if produit else 'N/A'}\n"
        f"Montant : {float(vente.montant_net or 0):,.2f} MAD\n\n"
        "Cette action est irreversible.",
    ):
        session.delete(vente)
        session.commit()
        charger_ventes(table)
        messagebox.showinfo("Succes", "Vente supprimee avec succes !")


# ══════════════════════════════════════════════════════════
#   IMPORTER VENTES
# ══════════════════════════════════════════════════════════
def _importer(table, format_):
    ext_map = {
        "excel": [("Fichier Excel", "*.xlsx"), ("Tous", "*.*")],
        "csv":   [("Fichier CSV",   "*.csv"),  ("Tous", "*.*")],
        "json":  [("Fichier JSON",  "*.json"), ("Tous", "*.*")],
    }
    chemin = filedialog.askopenfilename(
        title=f"Importer ventes depuis {format_.upper()}",
        filetypes=ext_map.get(format_, [("Tous", "*.*")]),
    )
    if not chemin:
        return

    try:
        if format_ == "excel":  lignes = _import_excel(chemin)
        elif format_ == "csv":  lignes = _import_csv(chemin)
        elif format_ == "json": lignes = _import_json(chemin)
        else:                   lignes = []

        if not lignes:
            messagebox.showwarning("Vide", "Aucune ligne trouvee dans le fichier.")
            return

        _fenetre_preview(table, lignes, chemin)
    except Exception as ex:
        messagebox.showerror("Erreur import", f"Erreur :\n{ex}")


def _import_excel(chemin):
    if not EXCEL_OK:
        messagebox.showerror("Manquant", "pip install openpyxl"); return []
    wb   = openpyxl.load_workbook(chemin, data_only=True)
    ws   = wb.active
    rows = list(ws.iter_rows(values_only=True))
    start = 1 if (rows and rows[0] and
                  any(str(c or "").lower() in ("client", "nom", "vente")
                      for c in rows[0])) else 0
    return [[str(c) if c is not None else "" for c in row]
            for row in rows[start:] if any(c is not None for c in row)]


def _import_csv(chemin):
    lignes = []
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        next(reader, None)
        for row in reader:
            if row:
                lignes.append(row)
    return lignes


def _import_json(chemin):
    with open(chemin, encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        data = data.get("ventes", data.get("sales", []))
    champs_j = ["client", "produit", "date_vente",
                "quantite", "prix", "reduction", "montant_net", "ville"]
    return [[str(item.get(k, "")) for k in champs_j]
            for item in data if isinstance(item, dict)]


def _fenetre_preview(table, lignes, chemin):
    t       = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")
    win = tk.Toplevel()
    win.title(f"Apercu import — {len(lignes)} ligne(s)")
    win.geometry("900x460")
    win.configure(bg=t["bg"])
    win.grab_set()

    hdr = tk.Frame(win, bg="#1565C0", height=50)
    hdr.pack(fill="x")
    tk.Label(hdr, text=f"📥  Apercu Import — {os.path.basename(chemin)}",
             bg="#1565C0", fg="white",
             font=("Arial", 13, "bold")).pack(pady=12)

    frame_prev = tk.Frame(win, bg=t["bg"])
    frame_prev.pack(fill="both", expand=True, padx=16, pady=10)

    cols_prev = ["Client", "Produit", "Date", "Qte",
                 "Prix", "Reduction", "Montant Net", "Ville"]

    prev_tree = ttk.Treeview(frame_prev, columns=cols_prev,
                             show="headings", height=12,
                             style="VN.Treeview")
    for col in cols_prev:
        prev_tree.heading(col, text=col)
        prev_tree.column(col, width=100, anchor="center")

    bg_row1 = "#1A1A35" if is_dark else "#FFFFFF"
    bg_row2 = "#12122A" if is_dark else "#FFF8F0"
    prev_tree.tag_configure("pair",   background=bg_row1)
    prev_tree.tag_configure("impair", background=bg_row2)

    for i, row in enumerate(lignes[:50]):
        # Sauter l'ID si present (9 colonnes) sinon prendre 8
        if len(row) >= 9:
            vals = row[1:9]
        elif len(row) >= 8:
            vals = row[:8]
        else:
            vals = list(row) + [""] * (8 - len(row))

        prev_tree.insert("", "end",
                         tags=("pair" if i % 2 == 0 else "impair",),
                         values=vals[:8])

    sb = ttk.Scrollbar(frame_prev, orient="vertical", command=prev_tree.yview)
    prev_tree.configure(yscroll=sb.set)
    sb.pack(side="right", fill="y")
    prev_tree.pack(fill="both", expand=True)

    if len(lignes) > 50:
        tk.Label(win, text=f"... et {len(lignes)-50} autres lignes.",
                 bg=t["bg"], fg="#FF8F00",
                 font=("Arial", 9, "italic")).pack()

    tk.Label(win,
             text=f"{len(lignes)} ligne(s) detectee(s).  "
                  "Seules les lignes avec client et produit existants sont importees.",
             bg=t["bg"], fg=t["text"],
             font=("Arial", 9)).pack(pady=(0, 4))

    btn_frame = tk.Frame(win, bg=t["bg"])
    btn_frame.pack(pady=8)

    def confirmer():
        nb_ok, nb_err = _inserer_lignes(lignes)
        charger_ventes(table)
        win.destroy()
        messagebox.showinfo(
            "Import termine",
            f"Import termine.\n\n"
            f"Inseres avec succes : {nb_ok}\n"
            f"Ignores (erreurs)   : {nb_err}",
        )

    _btn_ok(btn_frame, f"✅  Importer {len(lignes)} vente(s)",
            confirmer, "#1565C0", "#0D47A1").pack(side="left", padx=8)
    tk.Button(btn_frame, text="Annuler", command=win.destroy,
              bg=t["card"], fg=t["text"], font=("Arial", 11),
              relief="flat", cursor="hand2",
              padx=16, pady=9).pack(side="left", padx=8)


def _inserer_lignes(lignes):
    nb_ok = nb_err = 0
    for row in lignes:
        try:
            # Adapter : 9 colonnes (avec ID) ou 8 (sans ID)
            if len(row) >= 9:
                _, nom_c, nom_p, dt, qte, prix, red, net, ville = (
                    row[0], row[1], row[2], row[3], row[4],
                    row[5], row[6], row[7], row[8])
            elif len(row) >= 8:
                nom_c, nom_p, dt, qte, prix, red, net, ville = row[:8]
            else:
                nb_err += 1; continue

            client  = session.query(Client).filter_by(
                nom=str(nom_c).strip()).first()
            produit = session.query(Produit).filter_by(
                nom=str(nom_p).strip()).first()

            if not client or not produit:
                nb_err += 1; continue

            qte_i  = int(float(str(qte).replace(",", "") or 1))
            prix_f = float(str(prix).replace(",", "") or 0)
            red_f  = float(str(red).replace(",", "").replace("%", "") or 0)
            net_f  = float(str(net).replace(",", "") or 0)
            if net_f == 0 and prix_f > 0:
                net_f = round(prix_f * qte_i * (1 - red_f / 100), 2)

            from datetime import datetime
            try:
                date_v = datetime.strptime(str(dt).strip(), "%Y-%m-%d").date()
            except Exception:
                date_v = date.today()

            v = Vente(
                client_id=client.id,
                produit_id=produit.id,
                date_vente=date_v,
                quantite=qte_i,
                prix=prix_f,
                reduction=red_f,
                montant_net=net_f,
                ville=str(ville).strip() or "",
            )
            session.add(v)
            nb_ok += 1
        except Exception:
            nb_err += 1

    session.commit()
    return nb_ok, nb_err