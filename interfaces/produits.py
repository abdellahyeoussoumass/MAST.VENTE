import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from database.db import session
from database.models import Produit
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
    import csv
    CSV_OK = True
except ImportError:
    CSV_OK = False

try:
    import json
    JSON_OK = True
except ImportError:
    JSON_OK = False


# ══════════════════════════════════════════════════════════
#   CONSTANTES
# ══════════════════════════════════════════════════════════
EN_TETES = ["ID", "Reference", "Nom", "Categorie",
            "Prix HT", "TVA (%)", "Prix TTC", "Quantite"]
CHAMPS   = ["id", "reference", "nom", "categorie",
            "prix_ht", "tva", "prix_ttc", "quantite"]


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

    style.configure("PR.Treeview",
                    background=bg_tree, foreground=fg_text,
                    rowheight=32, fieldbackground=bg_tree,
                    borderwidth=0, font=("Arial", 10))
    style.configure("PR.Treeview.Heading",
                    background=bg_head, foreground="white",
                    font=("Arial", 10, "bold"),
                    relief="flat", borderwidth=0)
    style.map("PR.Treeview",
              background=[("selected", bg_sel)],
              foreground=[("selected", "#FFFFFF")])
    style.map("PR.Treeview.Heading",
              background=[("active", "#BF360C")])
    return bg_row1, bg_row2


# ══════════════════════════════════════════════════════════
#   HELPERS UI
# ══════════════════════════════════════════════════════════
def _fenetre_modale(titre, largeur=460, hauteur=560):
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
#   AFFICHER PRODUITS
# ══════════════════════════════════════════════════════════
def afficher_produits(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t       = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")

    # En-tete
    header = ctk.CTkFrame(parent, fg_color=t["card"], corner_radius=0)
    header.pack(fill="x")
    ctk.CTkLabel(header, text="📦  Gestion des Produits",
                 font=("Arial", 22, "bold"),
                 text_color="#E65100").pack(side="left", padx=24, pady=14)

    # Barre outils
    toolbar = ctk.CTkFrame(parent, fg_color="transparent")
    toolbar.pack(fill="x", padx=20, pady=(10, 4))

    btns = [
        ("➕  Ajouter",    "#E65100", "#BF360C", lambda: ajouter_produit(table)),
        ("✏️  Modifier",   "#FF8F00", "#E65100", lambda: modifier_produit(table)),
        ("🗑️  Supprimer",  "#C62828", "#8B0000", lambda: supprimer_produit(table)),
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

    colonnes = ("ID", "Reference", "Nom", "Categorie",
                "Prix HT", "TVA", "Prix TTC", "Qte")

    table = ttk.Treeview(frame_table, columns=colonnes,
                         show="headings", height=22,
                         style="PR.Treeview",
                         selectmode="browse")

    largeurs = {"ID": 45, "Reference": 110, "Nom": 180,
                "Categorie": 130, "Prix HT": 90,
                "TVA": 65, "Prix TTC": 90, "Qte": 70}

    for col in colonnes:
        table.heading(col, text=col,
                      command=lambda c=col: _trier(table, c))
        table.column(col, width=largeurs.get(col, 100), anchor="center")

    table.tag_configure("pair",   background=bg_row1)
    table.tag_configure("impair", background=bg_row2)

    # Couleur quantite faible
    table.tag_configure("stock_ok",     foreground="#43A047")
    table.tag_configure("stock_faible", foreground="#FFA000")
    table.tag_configure("stock_vide",   foreground="#E53935")

    sb_v = ttk.Scrollbar(frame_table, orient="vertical",   command=table.yview)
    sb_h = ttk.Scrollbar(frame_table, orient="horizontal", command=table.xview)
    table.configure(yscroll=sb_v.set, xscroll=sb_h.set)
    sb_v.pack(side="right",  fill="y")
    sb_h.pack(side="bottom", fill="x")
    table.pack(fill="both", expand=True, padx=2, pady=2)

    table.bind("<Double-1>", lambda e: modifier_produit(table))

    charger_produits(table, count_var)
    return table


# ══════════════════════════════════════════════════════════
#   CHARGER PRODUITS
# ══════════════════════════════════════════════════════════
def charger_produits(table, count_var=None):
    for row in table.get_children():
        table.delete(row)

    produits = session.query(Produit).order_by(Produit.id.desc()).all()

    for i, p in enumerate(produits):
        qte  = p.quantite or 0
        tags = ["pair" if i % 2 == 0 else "impair"]
        if qte == 0:
            tags.append("stock_vide")
        elif qte <= 5:
            tags.append("stock_faible")
        else:
            tags.append("stock_ok")

        table.insert("", "end", tags=tuple(tags), values=(
            p.id,
            p.reference or "-",
            p.nom,
            p.categorie or "-",
            f"{float(p.prix_ht):,.2f}",
            f"{float(p.tva):.0f}%",
            f"{float(p.prix_ttc):,.2f}",
            qte,
        ))

    if count_var is not None:
        count_var.set(f"{len(produits)} produit(s)")


# ══════════════════════════════════════════════════════════
#   RECHERCHER
# ══════════════════════════════════════════════════════════
def rechercher(texte, table):
    for row in table.get_children():
        table.delete(row)

    produits = session.query(Produit).filter(
        Produit.nom.ilike(f"%{texte}%")
    ).all()

    for i, p in enumerate(produits):
        qte  = p.quantite or 0
        tags = ["pair" if i % 2 == 0 else "impair"]
        if qte == 0:
            tags.append("stock_vide")
        elif qte <= 5:
            tags.append("stock_faible")
        else:
            tags.append("stock_ok")

        table.insert("", "end", tags=tuple(tags), values=(
            p.id,
            p.reference or "-",
            p.nom,
            p.categorie or "-",
            f"{float(p.prix_ht):,.2f}",
            f"{float(p.tva):.0f}%",
            f"{float(p.prix_ttc):,.2f}",
            qte,
        ))


# ══════════════════════════════════════════════════════════
#   AJOUTER PRODUIT
# ══════════════════════════════════════════════════════════
def ajouter_produit(table):
    win, body, t = _fenetre_modale("➕  Nouveau Produit", 460, 580)

    _lbl(body, "Reference", t)
    e_ref = _entry(body, t)

    _lbl(body, "Nom du produit", t)
    e_nom = _entry(body, t)

    _lbl(body, "Categorie", t)
    e_cat = _entry(body, t)

    # Prix HT + TVA sur la meme ligne
    prix_row = tk.Frame(body, bg=t["bg"])
    prix_row.pack(fill="x", pady=(10, 0))

    col1 = tk.Frame(prix_row, bg=t["bg"])
    col1.pack(side="left", expand=True, fill="x", padx=(0, 8))
    tk.Label(col1, text="Prix HT (MAD)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_ht = tk.Entry(col1, font=("Arial", 11),
                    bg=t["card"], fg=t["text"],
                    insertbackground=t["text"],
                    relief="flat", bd=6)
    e_ht.pack(fill="x", ipady=4)

    col2 = tk.Frame(prix_row, bg=t["bg"])
    col2.pack(side="left", expand=True, fill="x")
    tk.Label(col2, text="TVA (%)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_tva = tk.Entry(col2, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6)
    e_tva.insert(0, "20")
    e_tva.pack(fill="x", ipady=4)

    # Preview TTC
    lbl_ttc = tk.Label(body, text="",
                       bg=t["bg"], fg="#FF8F00",
                       font=("Arial", 11, "bold"), anchor="w")
    lbl_ttc.pack(fill="x", pady=(6, 0))

    def maj_ttc(*_):
        try:
            ht  = float(e_ht.get())
            tva = float(e_tva.get())
            ttc = round(ht * (1 + tva / 100), 2)
            lbl_ttc.config(text=f"Prix TTC calculé : {ttc:,.2f} MAD")
        except ValueError:
            lbl_ttc.config(text="")

    e_ht.bind("<KeyRelease>",  maj_ttc)
    e_tva.bind("<KeyRelease>", maj_ttc)

    _lbl(body, "Quantite en stock", t)
    e_qte = _entry(body, t, default="0")

    _sep(body, t)
    btn_frame = tk.Frame(body, bg=t["bg"])
    btn_frame.pack(fill="x", pady=(12, 4))

    def sauvegarder():
        try:
            ht  = float(e_ht.get())
            tva = float(e_tva.get())
            ttc = round(ht * (1 + tva / 100), 2)

            if not e_nom.get().strip():
                messagebox.showerror("Erreur", "Le nom est obligatoire.", parent=win)
                return

            p = Produit(
                reference=e_ref.get().strip(),
                nom=e_nom.get().strip(),
                categorie=e_cat.get().strip(),
                prix_ht=ht, tva=tva,
                prix_ttc=ttc,
                quantite=int(e_qte.get() or 0),
            )
            session.add(p)
            session.commit()
            charger_produits(table)
            win.destroy()
            messagebox.showinfo("Succes", "Produit ajoute avec succes !")
        except ValueError:
            messagebox.showerror("Erreur", "Verifiez les valeurs numeriques.", parent=win)

    _btn_ok(btn_frame, "💾  Sauvegarder", sauvegarder).pack(
        side="left", expand=True, fill="x", padx=(0, 6))
    tk.Button(btn_frame, text="Annuler", command=win.destroy,
              bg=t["card"], fg=t["text"], font=("Arial", 11),
              relief="flat", cursor="hand2", pady=9).pack(
              side="left", expand=True, fill="x")


# ══════════════════════════════════════════════════════════
#   MODIFIER PRODUIT
# ══════════════════════════════════════════════════════════
def modifier_produit(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez un produit a modifier !")
        return

    valeurs = table.item(sel[0])["values"]
    produit = session.query(Produit).filter_by(id=valeurs[0]).first()
    if not produit:
        return

    win, body, t = _fenetre_modale("✏️  Modifier le Produit", 460, 580)

    _lbl(body, "Reference", t)
    e_ref = _entry(body, t, default=produit.reference or "")

    _lbl(body, "Nom du produit", t)
    e_nom = _entry(body, t, default=produit.nom or "")

    _lbl(body, "Categorie", t)
    e_cat = _entry(body, t, default=produit.categorie or "")

    # Prix HT + TVA cote a cote
    prix_row = tk.Frame(body, bg=t["bg"])
    prix_row.pack(fill="x", pady=(10, 0))

    col1 = tk.Frame(prix_row, bg=t["bg"])
    col1.pack(side="left", expand=True, fill="x", padx=(0, 8))
    tk.Label(col1, text="Prix HT (MAD)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_ht = tk.Entry(col1, font=("Arial", 11),
                    bg=t["card"], fg=t["text"],
                    insertbackground=t["text"],
                    relief="flat", bd=6)
    e_ht.insert(0, str(produit.prix_ht))
    e_ht.pack(fill="x", ipady=4)

    col2 = tk.Frame(prix_row, bg=t["bg"])
    col2.pack(side="left", expand=True, fill="x")
    tk.Label(col2, text="TVA (%)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_tva = tk.Entry(col2, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6)
    e_tva.insert(0, str(produit.tva))
    e_tva.pack(fill="x", ipady=4)

    # Preview TTC
    lbl_ttc = tk.Label(body, text=f"Prix TTC actuel : {float(produit.prix_ttc):,.2f} MAD",
                       bg=t["bg"], fg="#FF8F00",
                       font=("Arial", 11, "bold"), anchor="w")
    lbl_ttc.pack(fill="x", pady=(6, 0))

    def maj_ttc(*_):
        try:
            ht  = float(e_ht.get())
            tva = float(e_tva.get())
            ttc = round(ht * (1 + tva / 100), 2)
            lbl_ttc.config(text=f"Prix TTC calcule : {ttc:,.2f} MAD")
        except ValueError:
            lbl_ttc.config(text="")

    e_ht.bind("<KeyRelease>",  maj_ttc)
    e_tva.bind("<KeyRelease>", maj_ttc)

    _lbl(body, "Quantite en stock", t)
    e_qte = _entry(body, t, default=str(produit.quantite or 0))

    _sep(body, t)
    btn_frame = tk.Frame(body, bg=t["bg"])
    btn_frame.pack(fill="x", pady=(12, 4))

    def sauvegarder():
        try:
            ht  = float(e_ht.get())
            tva = float(e_tva.get())
            ttc = round(ht * (1 + tva / 100), 2)

            if not e_nom.get().strip():
                messagebox.showerror("Erreur", "Le nom est obligatoire.", parent=win)
                return

            produit.reference = e_ref.get().strip()
            produit.nom       = e_nom.get().strip()
            produit.categorie = e_cat.get().strip()
            produit.prix_ht   = ht
            produit.tva       = tva
            produit.prix_ttc  = ttc
            produit.quantite  = int(e_qte.get() or 0)
            session.commit()
            charger_produits(table)
            win.destroy()
            messagebox.showinfo("Succes", "Produit modifie avec succes !")
        except ValueError:
            messagebox.showerror("Erreur", "Verifiez les valeurs numeriques.", parent=win)

    _btn_ok(btn_frame, "💾  Sauvegarder", sauvegarder).pack(
        side="left", expand=True, fill="x", padx=(0, 6))
    tk.Button(btn_frame, text="Annuler", command=win.destroy,
              bg=t["card"], fg=t["text"], font=("Arial", 11),
              relief="flat", cursor="hand2", pady=9).pack(
              side="left", expand=True, fill="x")


# ══════════════════════════════════════════════════════════
#   SUPPRIMER PRODUIT
# ══════════════════════════════════════════════════════════
def supprimer_produit(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez un produit a supprimer !")
        return

    valeurs = table.item(sel[0])["values"]
    produit = session.query(Produit).filter_by(id=valeurs[0]).first()
    if not produit:
        return

    if messagebox.askyesno(
        "Confirmation",
        f"Supprimer le produit : {produit.nom} ?\n"
        f"Reference : {produit.reference or '-'}\n"
        f"Prix TTC  : {float(produit.prix_ttc):,.2f} MAD\n\n"
        "Attention : les ventes liees a ce produit seront impactees.",
    ):
        session.delete(produit)
        session.commit()
        charger_produits(table)
        messagebox.showinfo("Succes", "Produit supprime avec succes !")


# ══════════════════════════════════════════════════════════
#   IMPORTER PRODUITS
# ══════════════════════════════════════════════════════════
def _importer(table, format_):
    ext_map = {
        "excel": [("Fichier Excel", "*.xlsx"), ("Tous", "*.*")],
        "csv":   [("Fichier CSV",   "*.csv"),  ("Tous", "*.*")],
        "json":  [("Fichier JSON",  "*.json"), ("Tous", "*.*")],
    }
    chemin = filedialog.askopenfilename(
        title=f"Importer depuis {format_.upper()}",
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
    # Sauter la ligne de titre si elle contient "produit" ou "nom"
    start = 1 if (rows and rows[0] and
                  any(str(c or "").lower() in ("nom", "produit", "reference")
                      for c in rows[0])) else 0
    return [[str(c) if c is not None else "" for c in row]
            for row in rows[start:] if any(c is not None for c in row)]


def _import_csv(chemin):
    import csv as csv_mod
    lignes = []
    with open(chemin, encoding="utf-8-sig", newline="") as f:
        reader = csv_mod.reader(f)
        next(reader, None)  # ignorer en-tete
        for row in reader:
            if row:
                lignes.append(row)
    return lignes


def _import_json(chemin):
    import json as json_mod
    with open(chemin, encoding="utf-8") as f:
        data = json_mod.load(f)
    if isinstance(data, dict):
        data = data.get("produits", data.get("products", []))
    champs_json = ["reference", "nom", "categorie",
                   "prix_ht", "tva", "prix_ttc", "quantite"]
    return [[str(item.get(k, "")) for k in champs_json]
            for item in data if isinstance(item, dict)]


def _fenetre_preview(table, lignes, chemin):
    t   = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")
    win = tk.Toplevel()
    win.title(f"Apercu import — {len(lignes)} ligne(s)")
    win.geometry("860x480")
    win.configure(bg=t["bg"])
    win.grab_set()

    hdr = tk.Frame(win, bg="#1565C0", height=50)
    hdr.pack(fill="x")
    tk.Label(hdr, text=f"📥  Apercu Import — {os.path.basename(chemin)}",
             bg="#1565C0", fg="white",
             font=("Arial", 13, "bold")).pack(pady=12)

    frame_prev = tk.Frame(win, bg=t["bg"])
    frame_prev.pack(fill="both", expand=True, padx=16, pady=10)

    cols_prev = ["Reference", "Nom", "Categorie",
                 "Prix HT", "TVA", "Prix TTC", "Quantite"]

    prev_tree = ttk.Treeview(frame_prev, columns=cols_prev,
                             show="headings", height=12,
                             style="PR.Treeview")
    for col in cols_prev:
        prev_tree.heading(col, text=col)
        prev_tree.column(col, width=105, anchor="center")

    bg_row1 = "#1A1A35" if is_dark else "#FFFFFF"
    bg_row2 = "#12122A" if is_dark else "#FFF8F0"
    prev_tree.tag_configure("pair",   background=bg_row1)
    prev_tree.tag_configure("impair", background=bg_row2)

    for i, row in enumerate(lignes[:50]):
        # Adapter selon la taille de la ligne (sans ID ou avec ID)
        if len(row) >= 8:
            vals = row[1:8]   # ignorer l'ID si present
        elif len(row) == 7:
            vals = row
        else:
            vals = list(row) + [""] * (7 - len(row))

        prev_tree.insert("", "end",
                         tags=("pair" if i % 2 == 0 else "impair",),
                         values=vals[:7])

    sb = ttk.Scrollbar(frame_prev, orient="vertical", command=prev_tree.yview)
    prev_tree.configure(yscroll=sb.set)
    sb.pack(side="right", fill="y")
    prev_tree.pack(fill="both", expand=True)

    if len(lignes) > 50:
        tk.Label(win, text=f"... et {len(lignes)-50} autres lignes.",
                 bg=t["bg"], fg="#FF8F00",
                 font=("Arial", 9, "italic")).pack()

    info_txt = (f"{len(lignes)} ligne(s) detectee(s).  "
                "Les doublons sur le nom sont ignores automatiquement.")
    tk.Label(win, text=info_txt,
             bg=t["bg"], fg=t["text"],
             font=("Arial", 9)).pack(pady=(0, 4))

    btn_frame = tk.Frame(win, bg=t["bg"])
    btn_frame.pack(pady=8)

    def confirmer():
        nb_ok, nb_err = _inserer_lignes(lignes)
        charger_produits(table)
        win.destroy()
        messagebox.showinfo(
            "Import termine",
            f"Import termine.\n\n"
            f"Inseres avec succes : {nb_ok}\n"
            f"Ignores (erreurs)   : {nb_err}",
        )

    _btn_ok(btn_frame, f"✅  Importer {len(lignes)} produit(s)",
            confirmer, "#1565C0", "#0D47A1").pack(side="left", padx=8)
    tk.Button(btn_frame, text="Annuler", command=win.destroy,
              bg=t["card"], fg=t["text"], font=("Arial", 11),
              relief="flat", cursor="hand2",
              padx=16, pady=9).pack(side="left", padx=8)


def _inserer_lignes(lignes):
    nb_ok = nb_err = 0
    for row in lignes:
        try:
            # Accepter 7 colonnes (sans ID) ou 8 (avec ID)
            if len(row) >= 8:
                _, ref, nom, cat, ht, tva, ttc, qte = (
                    row[0], row[1], row[2], row[3],
                    row[4], row[5], row[6], row[7])
            elif len(row) == 7:
                ref, nom, cat, ht, tva, ttc, qte = row
            else:
                nb_err += 1; continue

            nom = str(nom).strip()
            if not nom:
                nb_err += 1; continue

            # Eviter les doublons sur le nom
            if session.query(Produit).filter_by(nom=nom).first():
                nb_err += 1; continue

            ht_f  = float(str(ht).replace(",", "") or 0)
            tva_f = float(str(tva).replace("%", "").replace(",", "") or 0)
            ttc_f = float(str(ttc).replace(",", "") or 0)
            if ttc_f == 0 and ht_f > 0:
                ttc_f = round(ht_f * (1 + tva_f / 100), 2)
            qte_i = int(float(str(qte).replace(",", "") or 0))

            p = Produit(
                reference=str(ref).strip() or None,
                nom=nom,
                categorie=str(cat).strip() or None,
                prix_ht=ht_f,
                tva=tva_f,
                prix_ttc=ttc_f,
                quantite=qte_i,
            )
            session.add(p)
            nb_ok += 1
        except Exception:
            nb_err += 1
    session.commit()
    return nb_ok, nb_err