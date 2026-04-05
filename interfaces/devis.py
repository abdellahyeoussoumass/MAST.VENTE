import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from database.db import session
from database.models import Devis, Client, Produit
from datetime import date
from utils.theme import get_theme
import os

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                    Paragraph, Spacer, HRFlowable)
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    PDF_OK = True
except ImportError:
    PDF_OK = False


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

    style.configure("Pro.Treeview",
                    background=bg_tree, foreground=fg_text,
                    rowheight=32, fieldbackground=bg_tree,
                    borderwidth=0, font=("Arial", 10))
    style.configure("Pro.Treeview.Heading",
                    background=bg_head, foreground="white",
                    font=("Arial", 10, "bold"),
                    relief="flat", borderwidth=0)
    style.map("Pro.Treeview",
              background=[("selected", bg_sel)],
              foreground=[("selected", "#FFFFFF")])
    style.map("Pro.Treeview.Heading",
              background=[("active", "#BF360C")])
    return bg_row1, bg_row2


# ══════════════════════════════════════════════════════════
#   CONSTANTES STATUT
# ══════════════════════════════════════════════════════════
STATUT_ICONE = {
    "Brouillon": "📝",
    "Envoye":    "📤",
    "Accepte":   "✅",
    "Refuse":    "❌",
}

STATUT_COULEUR = {
    "Brouillon": "#9E9E9E",
    "Envoye":    "#42A5F5",
    "Accepte":   "#43A047",
    "Refuse":    "#E53935",
}


# ══════════════════════════════════════════════════════════
#   HELPERS UI
# ══════════════════════════════════════════════════════════
def _fenetre_modale(titre, largeur=460, hauteur=600):
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
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(10, 2))


def _entry(body, t, default=""):
    e = tk.Entry(body, font=("Arial", 11),
                 bg=t["card"], fg=t["text"],
                 insertbackground=t["text"],
                 relief="flat", bd=6)
    e.pack(fill="x", ipady=4)
    if default:
        e.insert(0, default)
    return e


def _combo(body, t, values, default="", readonly=True):
    var = tk.StringVar(value=default)
    state = "readonly" if readonly else "normal"
    cb  = ttk.Combobox(body, textvariable=var, values=values,
                       font=("Arial", 11), state=state)
    cb.pack(fill="x", ipady=4)
    cb.var = var
    return cb


def _btn_ok(parent, text, cmd, t):
    return tk.Button(
        parent, text=text, command=cmd,
        bg="#E65100", fg="white",
        font=("Arial", 12, "bold"),
        relief="flat", cursor="hand2",
        padx=20, pady=10,
        activebackground="#BF360C",
        activeforeground="white",
    )


def _btn_cancel(parent, cmd, t):
    return tk.Button(
        parent, text="Annuler", command=cmd,
        bg=t["card"], fg=t["text"],
        font=("Arial", 11),
        relief="flat", cursor="hand2", pady=8,
    )


def _resume_frame(body, t, champs):
    info_bg = "#1A1A35" if t["bg"] in ("#1A1A2E", "#0F0F1A") else "#FFF0E5"
    frame   = tk.Frame(body, bg=info_bg)
    frame.pack(fill="x", pady=(0, 12))
    for label, val in champs:
        row = tk.Frame(frame, bg=info_bg)
        row.pack(fill="x", padx=14, pady=3)
        tk.Label(row, text=f"{label} :", bg=info_bg,
                 fg="#FF8F00", font=("Arial", 10, "bold"),
                 width=11, anchor="w").pack(side="left")
        tk.Label(row, text=str(val), bg=info_bg,
                 fg=t["text"], font=("Arial", 10)).pack(side="left")


# ══════════════════════════════════════════════════════════
#   TRI COLONNES
# ══════════════════════════════════════════════════════════
_tri_etat = {}

def _trier(table, col):
    rows = [(table.set(r, col), r) for r in table.get_children("")]
    rev  = _tri_etat.get(col, False)
    try:
        rows.sort(key=lambda x: float(x[0].replace(",", "").replace("%", "")), reverse=rev)
    except ValueError:
        rows.sort(key=lambda x: x[0].lower(), reverse=rev)
    for i, (_, r) in enumerate(rows):
        table.move(r, "", i)
    _tri_etat[col] = not rev


# ══════════════════════════════════════════════════════════
#   AFFICHER DEVIS
# ══════════════════════════════════════════════════════════
def afficher_devis(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t       = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")

    # En-tete
    header = ctk.CTkFrame(parent, fg_color=t["card"], corner_radius=0)
    header.pack(fill="x")
    ctk.CTkLabel(header, text="📋  Gestion des Devis",
                 font=("Arial", 22, "bold"),
                 text_color="#E65100").pack(side="left", padx=24, pady=14)

    # Barre outils
    toolbar = ctk.CTkFrame(parent, fg_color="transparent")
    toolbar.pack(fill="x", padx=20, pady=(10, 4))

    btns = [
        ("➕  Ajouter",    "#E65100", "#BF360C", lambda: ajouter_devis(table)),
        ("✏️  Modifier",   "#FF8F00", "#E65100", lambda: modifier_statut(table)),
        ("🗑️  Supprimer",  "#C62828", "#8B0000", lambda: supprimer_devis(table)),
        ("👁️  Consulter",  "#2E7D32", "#1B5E20", lambda: detail_devis(table)),
        ("📤  Exporter",   "#1565C0", "#0D47A1", lambda: exporter_devis(table)),
    ]
    for txt, fg, hover, cmd in btns:
        ctk.CTkButton(toolbar, text=txt, width=126, height=36,
                      fg_color=fg, hover_color=hover,
                      font=("Arial", 11, "bold"),
                      corner_radius=8,
                      command=cmd).pack(side="left", padx=4)

    # Compteur
    count_var = tk.StringVar()
    ctk.CTkLabel(toolbar, textvariable=count_var,
                 font=("Arial", 11),
                 text_color="#FF8F00").pack(side="right", padx=10)

    # ── Barre de recherche + filtre statut ───────────────────────────────────
    search_frame = ctk.CTkFrame(parent, fg_color="transparent")
    search_frame.pack(fill="x", padx=20, pady=(0, 4))

    # Etat interne du filtre
    filtre_actif  = {"valeur": None}
    dropdown_ref  = [None]

    FILTRES_STATUT = [
        ("Tous",      "#555555"),
        ("Brouillon", "#9E9E9E"),
        ("Envoye",    "#42A5F5"),
        ("Accepte",   "#43A047"),
        ("Refuse",    "#E53935"),
    ]
    ICONES_F = {
        "Tous": "📋", "Brouillon": "📝",
        "Envoye": "📤", "Accepte": "✅", "Refuse": "❌",
    }

    # Bouton filtre statut
    var_btn_txt = tk.StringVar(value="⚙  Filtre ▾")
    btn_filtre  = tk.Button(
        search_frame, textvariable=var_btn_txt,
        bg="#FF8F00", fg="white",
        font=("Arial", 10, "bold"),
        relief="flat", cursor="hand2",
        padx=10, pady=6,
        activebackground="#E65100",
        activeforeground="white",
    )
    btn_filtre.pack(side="left", padx=(0, 4))

    # Cadre du champ texte
    frame_input = tk.Frame(search_frame, bg=t["card"],
                           highlightbackground="#E65100",
                           highlightthickness=1)
    frame_input.pack(side="left")

    recherche_var = tk.StringVar()
    PLACEHOLDER   = "🔍  Rechercher dans tous les devis..."

    e_search = tk.Entry(
        frame_input, textvariable=recherche_var,
        font=("Arial", 11), bg=t["card"], fg="#888888",
        insertbackground=t["text"],
        relief="flat", bd=6, width=38,
    )
    e_search.insert(0, PLACEHOLDER)
    e_search.pack(side="left", ipady=5)

    btn_clear = tk.Button(
        frame_input, text=" ✕ ",
        bg=t["card"], fg="#C62828",
        font=("Arial", 10, "bold"),
        relief="flat", cursor="hand2",
        activebackground=t["card"],
        activeforeground="#8B0000", bd=0,
    )

    table_ref = [None]

    def _show_ph():
        if not recherche_var.get():
            e_search.configure(fg="#888888")
            e_search.delete(0, "end")
            e_search.insert(0, PLACEHOLDER)
    def _hide_ph(e=None):
        if e_search.get() == PLACEHOLDER:
            e_search.delete(0, "end")
            e_search.configure(fg=t["text"])
    e_search.bind("<FocusIn>",  _hide_ph)
    e_search.bind("<FocusOut>", lambda e: _show_ph() if not recherche_var.get() else None)

    def _effacer_tout():
        filtre_actif["valeur"] = None
        var_btn_txt.set("⚙  Filtre ▾")
        btn_filtre.configure(bg="#FF8F00")
        recherche_var.set("")
        _show_ph()
        btn_clear.pack_forget()
        if table_ref[0]: charger_devis(table_ref[0], count_var)

    btn_clear.configure(command=_effacer_tout)

    def _appliquer_recherche():
        if not table_ref[0]: return
        tbl = table_ref[0]
        texte = recherche_var.get().strip()
        if texte == PLACEHOLDER: texte = ""
        texte_low = texte.lower()
        statut_v  = filtre_actif["valeur"]

        devis_all = session.query(Devis).order_by(Devis.id.desc()).all()

        if statut_v and statut_v != "Tous":
            devis_all = [d for d in devis_all if (d.statut or "") == statut_v]

        if texte_low:
            def _match(d):
                client  = session.query(Client).filter_by(id=d.client_id).first()
                produit = session.query(Produit).filter_by(id=d.produit_id).first()
                return any(texte_low in str(v).lower() for v in [
                    d.numero_devis, client.nom if client else "",
                    produit.nom if produit else "",
                    d.statut, d.prix_total, d.date_devis,
                ])
            devis_all = [d for d in devis_all if _match(d)]

        for row in tbl.get_children(): tbl.delete(row)
        for i, d in enumerate(devis_all):
            client  = session.query(Client).filter_by(id=d.client_id).first()
            produit = session.query(Produit).filter_by(id=d.produit_id).first()
            statut_raw = d.statut or ""
            icone      = STATUT_ICONE.get(statut_raw, "")
            tags = ("pair" if i%2==0 else "impair",)
            if statut_raw in STATUT_COULEUR: tags = tags + (statut_raw,)
            tbl.insert("", "end", tags=tags, values=(
                d.id, d.numero_devis or "-",
                client.nom if client else "N/A",
                produit.nom if produit else "N/A",
                d.quantite,
                f"{d.prix_ht:,.2f}", f"{d.tva}%",
                f"{d.prix_ttc:,.2f}", f"{d.prix_total:,.2f}",
                f"{icone} {statut_raw}".strip() if icone else statut_raw,
                d.date_devis,
            ))

        nb = len(devis_all)
        suf = f"  |  Filtre: {statut_v}" if statut_v and statut_v != "Tous" else ""
        if texte: suf += f'  |  "{texte}"'
        count_var.set(f"{nb} devis{suf}")
        if texte or (statut_v and statut_v != "Tous"):
            btn_clear.pack(side="right")
        else:
            btn_clear.pack_forget()

    def _on_frappe(*_):
        texte = recherche_var.get()
        if texte and texte != PLACEHOLDER: _appliquer_recherche()
        elif not texte:
            btn_clear.pack_forget()
            if table_ref[0]: charger_devis(table_ref[0], count_var)

    recherche_var.trace("w", _on_frappe)

    def _fermer_dd():
        try:
            if dropdown_ref[0] and dropdown_ref[0].winfo_exists():
                dropdown_ref[0].destroy()
        except Exception: pass
        dropdown_ref[0] = None

    def _activer_filtre(valeur):
        _fermer_dd()
        filtre_actif["valeur"] = valeur
        if valeur == "Tous":
            var_btn_txt.set("⚙  Filtre ▾")
            btn_filtre.configure(bg="#FF8F00")
        else:
            icone = ICONES_F.get(valeur, "🔍")
            var_btn_txt.set(f"{icone} {valeur} ▾")
            col = {"Brouillon":"#9E9E9E","Envoye":"#1565C0",
                   "Accepte":"#2E7D32","Refuse":"#C62828"}.get(valeur,"#FF8F00")
            btn_filtre.configure(bg=col)
        _appliquer_recherche()

    def _ouvrir_filtre(e=None):
        _fermer_dd()
        is_d  = t["bg"] in ("#1A1A2E","#0F0F1A","#0D0D1A")
        bg_m  = "#1A1A35" if is_d else "#FFFFFF"
        bg_h  = "#252545" if is_d else "#FFF0E5"
        brd   = "#2A2A4A" if is_d else "#E8D8C8"
        muted = "#A0A0C0" if is_d else "#999999"

        dd = tk.Toplevel()
        dd.overrideredirect(True)
        dd.configure(bg=brd)
        dd.attributes("-topmost", True)
        dropdown_ref[0] = dd

        btn_filtre.update_idletasks()
        x = btn_filtre.winfo_rootx()
        y = btn_filtre.winfo_rooty() + btn_filtre.winfo_height() + 2
        h = len(FILTRES_STATUT)*38 + 28
        dd.geometry(f"210x{h}+{x}+{y}")

        inner = tk.Frame(dd, bg=bg_m)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        tk.Label(inner, text="  Filtrer par statut",
                 bg=bg_m, fg=muted,
                 font=("Arial",8,"bold"), anchor="w").pack(fill="x", padx=8, pady=(5,2))

        for label, couleur in FILTRES_STATUT:
            est_actif = filtre_actif["valeur"] == label
            bg_i = bg_h if est_actif else bg_m
            rf = tk.Frame(inner, bg=bg_i, cursor="hand2")
            rf.pack(fill="x", padx=4, pady=1)
            cv = tk.Canvas(rf, width=10, height=10, bg=bg_i, highlightthickness=0)
            cv.create_oval(1,1,9,9,fill=couleur,outline="")
            cv.pack(side="left", padx=(10,4), pady=10)
            tk.Label(rf, text="✓" if est_actif else "  ",
                     bg=bg_i, fg=couleur,
                     font=("Arial",10,"bold"), width=2).pack(side="left")
            lb = tk.Label(rf, text=label,
                          bg=bg_i, fg=couleur,
                          font=("Arial",10,"bold"), anchor="w", pady=8)
            lb.pack(side="left", fill="x", expand=True)

            def _mk(fr=rf, lb_=lb, cv_=cv, act=est_actif):
                def _in(e):
                    fr.configure(bg=bg_h); lb_.configure(bg=bg_h); cv_.configure(bg=bg_h)
                def _out(e):
                    c=bg_h if act else bg_m
                    fr.configure(bg=c); lb_.configure(bg=c); cv_.configure(bg=c)
                return _in, _out
            _in,_out=_mk()
            for w in (rf,lb,cv):
                w.bind("<Enter>",_in); w.bind("<Leave>",_out)
                w.bind("<Button-1>",lambda e,v=label: _activer_filtre(v))

        dd.bind("<FocusOut>", lambda e: _fermer_dd())
        dd.focus_set()

    btn_filtre.configure(command=_ouvrir_filtre)

    # Tableau
    frame_table = ctk.CTkFrame(parent, corner_radius=12)
    frame_table.pack(fill="both", expand=True, padx=20, pady=(4, 16))

    bg_row1, bg_row2 = _style_treeview(is_dark)

    colonnes = ("ID", "N Devis", "Client", "Produit",
                "Qte", "Prix HT", "TVA", "Prix TTC",
                "Total", "Statut", "Date")

    table = ttk.Treeview(frame_table, columns=colonnes,
                         show="headings", height=22,
                         style="Pro.Treeview",
                         selectmode="browse")

    largeurs = {"ID": 45, "N Devis": 100, "Client": 140,
                "Produit": 130, "Qte": 55, "Prix HT": 85,
                "TVA": 55, "Prix TTC": 90, "Total": 100,
                "Statut": 95, "Date": 100}

    for col in colonnes:
        table.heading(col, text=col,
                      command=lambda c=col: _trier(table, c))
        table.column(col, width=largeurs.get(col, 90), anchor="center")

    table.tag_configure("pair",    background=bg_row1)
    table.tag_configure("impair",  background=bg_row2)
    table.tag_configure("Accepte", foreground="#43A047")
    table.tag_configure("Refuse",  foreground="#E53935")
    table.tag_configure("Envoye",  foreground="#42A5F5")

    sb_v = ttk.Scrollbar(frame_table, orient="vertical",   command=table.yview)
    sb_h = ttk.Scrollbar(frame_table, orient="horizontal", command=table.xview)
    table.configure(yscroll=sb_v.set, xscroll=sb_h.set)
    sb_v.pack(side="right",  fill="y")
    sb_h.pack(side="bottom", fill="x")
    table.pack(fill="both", expand=True, padx=2, pady=2)

    table.bind("<Double-1>", lambda e: detail_devis(table))

    # Connecter la reference pour la barre de recherche
    table_ref[0] = table
    charger_devis(table, count_var)
    return table


# ══════════════════════════════════════════════════════════
#   CHARGER DEVIS
# ══════════════════════════════════════════════════════════
def charger_devis(table, count_var=None):
    for row in table.get_children():
        table.delete(row)

    devis_list = session.query(Devis).order_by(Devis.id.desc()).all()

    for i, d in enumerate(devis_list):
        client  = session.query(Client).filter_by(id=d.client_id).first()
        produit = session.query(Produit).filter_by(id=d.produit_id).first()

        statut_raw   = d.statut or ""
        statut_clean = statut_raw.replace("e", "e")
        icone        = STATUT_ICONE.get(statut_clean, "")

        tags = ("pair" if i % 2 == 0 else "impair",)
        if statut_clean in STATUT_COULEUR:
            tags = tags + (statut_clean,)

        table.insert("", "end", tags=tags, values=(
            d.id,
            d.numero_devis or "-",
            client.nom  if client  else "N/A",
            produit.nom if produit else "N/A",
            d.quantite,
            f"{d.prix_ht:,.2f}",
            f"{d.tva}%",
            f"{d.prix_ttc:,.2f}",
            f"{d.prix_total:,.2f}",
            f"{icone} {statut_raw}".strip() if icone else statut_raw,
            d.date_devis,
        ))

    if count_var is not None:
        count_var.set(f"{len(devis_list)} devis")


# ══════════════════════════════════════════════════════════
#   RECHERCHER
# ══════════════════════════════════════════════════════════
def rechercher(texte, table):
    for row in table.get_children():
        table.delete(row)

    texte_low  = texte.lower()
    devis_list = session.query(Devis).all()

    for i, d in enumerate(devis_list):
        client  = session.query(Client).filter_by(id=d.client_id).first()
        produit = session.query(Produit).filter_by(id=d.produit_id).first()
        nom_c   = client.nom  if client  else ""
        nom_p   = produit.nom if produit else ""

        if (texte_low in nom_c.lower()
                or texte_low in nom_p.lower()
                or texte_low in (d.numero_devis or "").lower()
                or texte_low in (d.statut or "").lower()):

            statut_raw   = d.statut or ""
            statut_clean = statut_raw.replace("e", "e")
            icone        = STATUT_ICONE.get(statut_clean, "")
            tags         = ("pair" if i % 2 == 0 else "impair",)

            table.insert("", "end", tags=tags, values=(
                d.id, d.numero_devis or "-",
                nom_c or "N/A", nom_p or "N/A",
                d.quantite,
                f"{d.prix_ht:,.2f}", f"{d.tva}%",
                f"{d.prix_ttc:,.2f}", f"{d.prix_total:,.2f}",
                f"{icone} {statut_raw}".strip() if icone else statut_raw,
                d.date_devis,
            ))


# ══════════════════════════════════════════════════════════
#   AJOUTER DEVIS
# ══════════════════════════════════════════════════════════
def ajouter_devis(table):
    clients  = session.query(Client).all()
    produits = session.query(Produit).all()

    win, body, t = _fenetre_modale("➕  Nouveau Devis", 460, 630)

    _lbl(body, "N Devis", t)
    e_num = _entry(body, t)

    _lbl(body, "Client", t)
    cb_client = _combo(body, t, [c.nom for c in clients])

    _lbl(body, "Produit", t)
    cb_produit = _combo(body, t, [p.nom for p in produits])

    _lbl(body, "Quantite", t)
    e_qte = _entry(body, t)

    # Prix HT + TVA cote a cote
    row_prix = tk.Frame(body, bg=t["bg"])
    row_prix.pack(fill="x", pady=(10, 0))

    col_ht = tk.Frame(row_prix, bg=t["bg"])
    col_ht.pack(side="left", expand=True, fill="x", padx=(0, 8))
    tk.Label(col_ht, text="Prix HT (MAD)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_ht = tk.Entry(col_ht, font=("Arial", 11), bg=t["card"], fg=t["text"],
                    insertbackground=t["text"], relief="flat", bd=6)
    e_ht.pack(fill="x", ipady=4)

    col_tva = tk.Frame(row_prix, bg=t["bg"])
    col_tva.pack(side="left", expand=True, fill="x")
    tk.Label(col_tva, text="TVA (%)", bg=t["bg"], fg=t["text"],
             font=("Arial", 11, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_tva = tk.Entry(col_tva, font=("Arial", 11), bg=t["card"], fg=t["text"],
                     insertbackground=t["text"], relief="flat", bd=6)
    e_tva.insert(0, "20")
    e_tva.pack(fill="x", ipady=4)

    _lbl(body, "Statut", t)
    cb_statut = _combo(body, t,
        ["Brouillon", "Envoye", "Accepte", "Refuse"],
        default="Brouillon")

    # Apercu prix en temps reel
    lbl_prev = tk.Label(body, text="", bg=t["bg"],
                        fg="#FF8F00", font=("Arial", 10, "bold"))
    lbl_prev.pack(fill="x", pady=(6, 0))

    def _get_prix_ht():
        """Retourne le prix HT : depuis champ manuel OU depuis le produit."""
        val = e_ht.get().strip()
        if val:
            return float(val)
        p = session.query(Produit).filter_by(nom=cb_produit.var.get()).first()
        if p:
            return float(p.prix_ht)
        return 0.0

    def _auto_fill_ht(*_):
        """Remplit le champ Prix HT quand on choisit un produit."""
        p = session.query(Produit).filter_by(nom=cb_produit.var.get()).first()
        if p:
            e_ht.delete(0, "end")
            e_ht.insert(0, str(p.prix_ht))
        maj_prev()

    def maj_prev(*_):
        try:
            ht  = _get_prix_ht()
            qte = int(e_qte.get() or 0)
            tva = float(e_tva.get() or 0)
            if ht > 0 and qte > 0:
                ttc = round(ht * (1 + tva / 100), 2)
                tot = round(ttc * qte, 2)
                lbl_prev.config(
                    text=f"HT: {ht:,.2f}  |  TTC: {ttc:,.2f}  |  Total: {tot:,.2f} MAD")
            else:
                lbl_prev.config(text="")
        except Exception:
            lbl_prev.config(text="")

    cb_produit.var.trace("w", _auto_fill_ht)
    e_ht.bind("<KeyRelease>",  maj_prev)
    e_qte.bind("<KeyRelease>", maj_prev)
    e_tva.bind("<KeyRelease>", maj_prev)

    def sauvegarder():
        try:
            client  = session.query(Client).filter_by(
                nom=cb_client.var.get()).first()
            produit = session.query(Produit).filter_by(
                nom=cb_produit.var.get()).first()
            if not client:
                messagebox.showerror("Erreur", "Selectionnez un client.", parent=win)
                return
            if not e_num.get().strip():
                messagebox.showerror("Erreur", "Le numero de devis est obligatoire.", parent=win)
                return
            qte   = int(e_qte.get())
            tva   = float(e_tva.get())
            ht    = _get_prix_ht()
            if ht <= 0:
                messagebox.showerror("Erreur", "Saisissez un prix HT.", parent=win)
                return
            ttc   = round(ht * (1 + tva / 100), 2)
            total = round(ttc * qte, 2)

            d = Devis(
                numero_devis=e_num.get().strip(),
                client_id=client.id,
                produit_id=produit.id if produit else None,
                categorie=produit.categorie if produit else None,
                prix_ht=ht, quantite=qte, tva=tva,
                prix_ttc=ttc, prix_total=total,
                statut=cb_statut.var.get(),
                date_devis=date.today(),
            )
            session.add(d)
            session.commit()
            charger_devis(table)
            win.destroy()
            messagebox.showinfo("Succes", "Devis ajoute avec succes !")
        except Exception as ex:
            messagebox.showerror("Erreur", f"Erreur : {ex}", parent=win)

    _btn_ok(body, "💾  Sauvegarder", sauvegarder, t).pack(fill="x", pady=(14, 4))
    _btn_cancel(body, win.destroy, t).pack(fill="x")


# ══════════════════════════════════════════════════════════
#   MODIFIER STATUT
# ══════════════════════════════════════════════════════════
def modifier_statut(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez un devis a modifier !")
        return

    valeurs = table.item(sel[0])["values"]
    devis   = session.query(Devis).filter_by(id=valeurs[0]).first()
    if not devis:
        return

    client  = session.query(Client).filter_by(id=devis.client_id).first()
    produit = session.query(Produit).filter_by(id=devis.produit_id).first()

    win, body, t = _fenetre_modale("✏️  Modifier le Devis", 420, 440)

    _resume_frame(body, t, [
        ("N Devis",  devis.numero_devis or "-"),
        ("Client",   client.nom  if client  else "N/A"),
        ("Produit",  produit.nom if produit else "N/A"),
        ("Total",    f"{devis.prix_total:,.2f} MAD"),
    ])

    _lbl(body, "Nouveau statut", t)
    cb = _combo(body, t,
        ["Brouillon", "Envoye", "Accepte", "Refuse"],
        default=devis.statut or "Brouillon")

    _lbl(body, "Note interne (optionnel)", t)
    e_note = tk.Text(body, height=3, font=("Arial", 10),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6, wrap="word")
    e_note.pack(fill="x")

    def sauvegarder():
        devis.statut = cb.var.get()
        session.commit()
        charger_devis(table)
        win.destroy()
        messagebox.showinfo("Succes", "Statut modifie avec succes !")

    _btn_ok(body, "💾  Sauvegarder", sauvegarder, t).pack(fill="x", pady=(14, 4))
    _btn_cancel(body, win.destroy, t).pack(fill="x")


# ══════════════════════════════════════════════════════════
#   SUPPRIMER DEVIS
# ══════════════════════════════════════════════════════════
def supprimer_devis(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez un devis a supprimer !")
        return

    valeurs = table.item(sel[0])["values"]
    devis   = session.query(Devis).filter_by(id=valeurs[0]).first()
    if not devis:
        return

    client = session.query(Client).filter_by(id=devis.client_id).first()
    nom_c  = client.nom if client else "N/A"

    if messagebox.askyesno(
        "Confirmation",
        f"Supprimer le devis {devis.numero_devis or devis.id} ?\n"
        f"Client : {nom_c}\n"
        f"Total  : {devis.prix_total:,.2f} MAD\n\n"
        "Cette action est irreversible.",
    ):
        session.delete(devis)
        session.commit()
        charger_devis(table)
        messagebox.showinfo("Succes", "Devis supprime avec succes !")


# ══════════════════════════════════════════════════════════
#   EXPORTER DEVIS (PDF au format professionnel)
# ══════════════════════════════════════════════════════════
def exporter_devis(table):
    sel = table.selection()
    if not sel:
        messagebox.showwarning("Attention", "Selectionnez un devis a exporter !")
        return

    valeurs = table.item(sel[0])["values"]
    devis   = session.query(Devis).filter_by(id=valeurs[0]).first()
    if not devis:
        return

    client  = session.query(Client).filter_by(id=devis.client_id).first()
    produit = session.query(Produit).filter_by(id=devis.produit_id).first()

    chemin = filedialog.asksaveasfilename(
        title="Exporter le devis en PDF",
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf"), ("Tous", "*.*")],
        initialfile=f"devis_{devis.numero_devis or devis.id}.pdf",
    )
    if not chemin:
        return

    try:
        _generer_pdf_devis(devis, client, produit, chemin)
        if messagebox.askyesno("Export reussi",
                               f"Devis exporte :\n{chemin}\n\nOuvrir maintenant ?"):
            os.startfile(chemin)
    except Exception as ex:
        messagebox.showerror("Erreur export", f"Erreur :\n{ex}")


def _generer_pdf_devis(devis, client, produit, chemin):
    """Genere un PDF professionnel au format devis (inspire de l'image exemple)."""
    if not PDF_OK:
        raise RuntimeError(
            "reportlab non installe.\nExecutez : pip install reportlab"
        )

    doc    = SimpleDocTemplate(chemin, pagesize=A4,
                               leftMargin=1.8*cm, rightMargin=1.8*cm,
                               topMargin=1.8*cm, bottomMargin=1.8*cm)
    styles = getSampleStyleSheet()
    story  = []

    ORANGE  = colors.HexColor("#E65100")
    ORANGE2 = colors.HexColor("#FF8F00")
    GRAY    = colors.HexColor("#555555")
    LGRAY   = colors.HexColor("#F5F5F5")
    WHITE   = colors.white
    BLACK   = colors.black

    # ── En-tête : titre DEVIS + numéro ──────────────────────────
    titre_style = ParagraphStyle(
        "TitreDevis", parent=styles["Title"],
        fontSize=26, textColor=ORANGE,
        spaceAfter=2, alignment=TA_LEFT,
        fontName="Helvetica-Bold",
    )
    story.append(Paragraph(f"DEVIS n° {devis.numero_devis or devis.id}", titre_style))
    story.append(HRFlowable(width="100%", thickness=3,
                             color=ORANGE, spaceAfter=10))

    # ── Bloc entreprise + destinataire ──────────────────────────
    try:
        from database.models import ExerciceComptable
        ex = session.query(ExerciceComptable).first() if False else None
    except Exception:
        ex = None

    entreprise_nom = "VentePro"
    entreprise_adr = ""
    if client:
        client_ville = getattr(client, "ville", "") or ""
        client_adr   = getattr(client, "adresse", "") or ""
        client_email = getattr(client, "email", "") or ""
        client_tel   = getattr(client, "telephone", "") or ""
    else:
        client_ville = client_adr = client_email = client_tel = ""

    bloc_data = [[
        Paragraph(f"""<b>{entreprise_nom}</b><br/>
                      {entreprise_adr if entreprise_adr else "— Votre adresse —"}<br/>
                      """,
                  ParagraphStyle("Ent", parent=styles["Normal"],
                                 fontSize=9, textColor=BLACK)),
        Paragraph(f"""<b>Destinataire</b><br/>
                      <b>{client.nom if client else "N/A"}</b><br/>
                      {client_adr}<br/>
                      {client_ville}<br/>
                      {client_email}<br/>
                      {client_tel}
                      """,
                  ParagraphStyle("Cli", parent=styles["Normal"],
                                 fontSize=9, textColor=BLACK,
                                 alignment=TA_RIGHT)),
    ]]
    bloc_tbl = Table(bloc_data, colWidths=[9*cm, 8*cm])
    bloc_tbl.setStyle(TableStyle([
        ("VALIGN",  (0,0), (-1,-1), "TOP"),
        ("TOPPADDING",   (0,0), (-1,-1), 4),
        ("BOTTOMPADDING",(0,0), (-1,-1), 4),
    ]))
    story.append(bloc_tbl)
    story.append(Spacer(1, 0.6*cm))

    # ── Informations devis ──────────────────────────────────────
    info_data = [
        ["Date du devis :",       str(devis.date_devis),
         "Reference :",           devis.numero_devis or str(devis.id)],
        ["Validite :",            "30 jours",
         "Statut :",              devis.statut or "-"],
        ["Contact client :",      client.nom if client else "N/A",
         "Telephone :",           client_tel or "-"],
    ]
    info_tbl = Table(info_data, colWidths=[4*cm, 5*cm, 4*cm, 4.2*cm])
    info_tbl.setStyle(TableStyle([
        ("FONTNAME",      (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",      (0,0), (-1,-1), 8.5),
        ("FONTNAME",      (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME",      (2,0), (2,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (0,0), (0,-1), ORANGE),
        ("TEXTCOLOR",     (2,0), (2,-1), ORANGE),
        ("ROWBACKGROUNDS",(0,0), (-1,-1),
         [colors.HexColor("#FFF8F0"), WHITE]),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("GRID",          (0,0), (-1,-1), 0.3, colors.HexColor("#E8D8C8")),
    ]))
    story.append(info_tbl)
    story.append(Spacer(1, 0.8*cm))

    # ── Tableau de lignes du devis ──────────────────────────────
    col_hdrs = ["Description", "Quantite", "Unite",
                "Prix unitaire HT", "% TVA", "Total TVA", "Total TTC"]

    qte   = devis.quantite or 1
    ht_u  = float(devis.prix_ht or 0)
    tva_p = float(devis.tva or 20)
    ttc_u = round(ht_u * (1 + tva_p/100), 2)
    tva_u = round(ttc_u - ht_u, 2)
    tot_ht  = round(ht_u  * qte, 2)
    tot_tva = round(tva_u * qte, 2)
    tot_ttc = round(ttc_u * qte, 2)

    prod_nom = produit.nom if produit else (devis.categorie or "-")

    lines = [col_hdrs,
             [prod_nom, str(qte), "u",
              f"{ht_u:,.2f}",
              f"{tva_p:.0f} %",
              f"{tot_tva:,.2f}",
              f"{tot_ttc:,.2f}"],
             ["", "", "", "Total HT",  "",
              f"{tot_ht:,.2f}", ""],
             ["", "", "", "Total TVA", "",
              f"{tot_tva:,.2f}", ""],
             ["", "", "", "", "",
              "Total TTC", f"{tot_ttc:,.2f}"],
             ]

    col_w = [5.5*cm, 1.8*cm, 1.5*cm, 2.8*cm, 1.5*cm, 2.2*cm, 2.2*cm]
    data_tbl = Table(lines, colWidths=col_w, repeatRows=1)
    data_tbl.setStyle(TableStyle([
        # En-tete
        ("BACKGROUND",    (0,0), (-1,0), ORANGE),
        ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
        ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0), (-1,0), 8),
        ("ALIGN",         (1,0), (-1,-1), "RIGHT"),
        ("ALIGN",         (0,0), (0,-1), "LEFT"),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("TOPPADDING",    (0,0), (-1,0), 8),
        # Lignes donnees
        ("FONTSIZE",      (0,1), (-1,-1), 8.5),
        ("ROWBACKGROUNDS",(0,1), (-1,2),
         [WHITE, colors.HexColor("#FFF8F0")]),
        ("TOPPADDING",    (0,1), (-1,-1), 5),
        ("BOTTOMPADDING", (0,1), (-1,-1), 5),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
        ("RIGHTPADDING",  (0,0), (-1,-1), 6),
        ("GRID",          (0,0), (-1,2),  0.4, colors.HexColor("#E8D8C8")),
        # Lignes totaux
        ("FONTNAME",      (3,2), (-1,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (3,2), (-1,-1), ORANGE),
        ("LINEABOVE",     (0,2), (-1,2),  0.8, colors.HexColor("#E8D8C8")),
        # Ligne Total TTC (derniere)
        ("BACKGROUND",    (0,-1), (-1,-1), colors.HexColor("#BF360C")),
        ("TEXTCOLOR",     (0,-1), (-1,-1), WHITE),
        ("FONTNAME",      (0,-1), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",      (0,-1), (-1,-1), 10),
    ]))
    story.append(data_tbl)
    story.append(Spacer(1, 0.8*cm))

    # ── Signature / Bon pour accord ─────────────────────────────
    sig_data = [["Signature du client (precedee de la mention « Bon pour accord »):"]]
    sig_tbl  = Table(sig_data, colWidths=[17*cm])
    sig_tbl.setStyle(TableStyle([
        ("FONTNAME",      (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE",      (0,0), (-1,-1), 8.5),
        ("TEXTCOLOR",     (0,0), (-1,-1), GRAY),
        ("BACKGROUND",    (0,0), (-1,-1), LGRAY),
        ("TOPPADDING",    (0,0), (-1,-1), 40),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LEFTPADDING",   (0,0), (-1,-1), 10),
        ("BOX",           (0,0), (-1,-1), 0.5, colors.HexColor("#E8D8C8")),
    ]))
    story.append(sig_tbl)
    story.append(Spacer(1, 0.6*cm))

    # ── Pied de page ────────────────────────────────────────────
    story.append(HRFlowable(width="100%", thickness=1,
                             color=colors.HexColor("#E8D8C8"), spaceBefore=6))
    story.append(Paragraph(
        f"VentePro  —  Devis N° {devis.numero_devis or devis.id}  "
        f"—  {str(devis.date_devis)}  —  Valable 30 jours",
        ParagraphStyle("Footer", parent=styles["Normal"],
                       fontSize=7.5,
                       textColor=colors.HexColor("#999999"),
                       alignment=TA_CENTER)))
    doc.build(story)


# ══════════════════════════════════════════════════════════
#   DETAIL DEVIS  (double-clic)
# ══════════════════════════════════════════════════════════
def detail_devis(table):
    sel = table.selection()
    if not sel:
        return

    valeurs = table.item(sel[0])["values"]
    devis   = session.query(Devis).filter_by(id=valeurs[0]).first()
    if not devis:
        return

    client  = session.query(Client).filter_by(id=devis.client_id).first()
    produit = session.query(Produit).filter_by(id=devis.produit_id).first()

    win, body, t = _fenetre_modale(
        f"Detail — {devis.numero_devis or devis.id}", 420, 440)

    info_bg = "#1A1A35" if t["bg"] in ("#1A1A2E", "#0F0F1A") else "#FFF0E5"
    frame   = tk.Frame(body, bg=info_bg)
    frame.pack(fill="both", expand=True)

    champs = [
        ("N Devis",  devis.numero_devis or "-"),
        ("Date",     str(devis.date_devis)),
        ("Client",   client.nom  if client  else "N/A"),
        ("Produit",  produit.nom if produit else "N/A"),
        ("Quantite", str(devis.quantite)),
        ("Prix HT",  f"{devis.prix_ht:,.2f} MAD"),
        ("TVA",      f"{devis.tva}%"),
        ("Prix TTC", f"{devis.prix_ttc:,.2f} MAD"),
        ("Total",    f"{devis.prix_total:,.2f} MAD"),
        ("Statut",   devis.statut or "-"),
    ]

    for label, val in champs:
        row = tk.Frame(frame, bg=info_bg)
        row.pack(fill="x", padx=16, pady=5)
        tk.Label(row, text=f"{label} :", bg=info_bg,
                 fg="#FF8F00", font=("Arial", 11, "bold"),
                 width=11, anchor="w").pack(side="left")
        tk.Label(row, text=val, bg=info_bg,
                 fg=t["text"], font=("Arial", 11)).pack(side="left")
        tk.Frame(frame, bg="#333355" if t["bg"] in ("#1A1A2E",)
                 else "#E8D8C8", height=1).pack(fill="x", padx=16)

    tk.Button(body, text="Fermer", command=win.destroy,
              bg="#E65100", fg="white",
              font=("Arial", 12, "bold"),
              relief="flat", cursor="hand2", pady=10
              ).pack(fill="x", pady=(14, 0))