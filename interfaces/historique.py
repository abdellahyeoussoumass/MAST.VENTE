import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from database.db import session
from database.models import Vente, Client, Produit
from datetime import date
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from utils.theme import get_theme


# ══════════════════════════════════════════════════════════
#   CONSTANTES
# ══════════════════════════════════════════════════════════
PALETTE = {
    "orange":       "#E65100",
    "orange_light": "#FF8F00",
    "orange_gold":  "#FFB300",
    "green":        "#2E7D32",
    "green_light":  "#43A047",
    "teal":         "#00838F",
    "blue":         "#1565C0",
    "purple":       "#6A1B9A",
    "red":          "#C62828",
}

D_BG    = "#0D0D1A"
D_CARD  = "#12122A"
D_CARD2 = "#1A1A35"
D_GRID  = "#22223A"
D_TEXT  = "#E8E0FF"
D_MUTED = "#7070A0"

L_BG    = "#FFF8F0"
L_CARD  = "#FFFFFF"
L_CARD2 = "#FFF0E5"
L_GRID  = "#EEE5D8"
L_TEXT  = "#1A1A1A"
L_MUTED = "#777777"

MOIS = ["Jan", "Fev", "Mar", "Avr", "Mai", "Jun",
        "Jul", "Aou", "Sep", "Oct", "Nov", "Dec"]


# ══════════════════════════════════════════════════════════
#   CONFIGURATION MATPLOTLIB
# ══════════════════════════════════════════════════════════
def _cfg_mpl(is_dark):
    bg    = D_BG    if is_dark else L_BG
    card  = D_CARD  if is_dark else L_CARD
    grid  = D_GRID  if is_dark else L_GRID
    text  = D_TEXT  if is_dark else L_TEXT
    muted = D_MUTED if is_dark else L_MUTED
    plt.rcParams.update({
        "figure.facecolor":   bg,
        "axes.facecolor":     card,
        "axes.edgecolor":     grid,
        "axes.labelcolor":    muted,
        "axes.titlecolor":    text,
        "axes.titlepad":      12,
        "axes.grid":          True,
        "axes.grid.axis":     "y",
        "grid.color":         grid,
        "grid.linewidth":     0.5,
        "grid.linestyle":     "--",
        "grid.alpha":         0.6,
        "xtick.color":        muted,
        "ytick.color":        muted,
        "xtick.labelsize":    8.5,
        "ytick.labelsize":    8.5,
        "text.color":         text,
        "font.family":        "DejaVu Sans",
        "font.size":          9,
        "axes.spines.top":    False,
        "axes.spines.right":  False,
        "axes.spines.left":   False,
        "axes.spines.bottom": True,
        "figure.dpi":         110,
        "legend.frameon":     False,
    })


# ══════════════════════════════════════════════════════════
#   HELPERS
# ══════════════════════════════════════════════════════════
def _fmt(val, _=None):
    if val >= 1_000_000:
        return f"{val/1_000_000:.1f}M"
    if val >= 1_000:
        return f"{val/1_000:.0f}k"
    return f"{int(val)}"


def _ax_style(ax, is_dark):
    bg   = D_CARD if is_dark else L_CARD
    grid = D_GRID if is_dark else L_GRID
    ax.set_facecolor(bg)
    ax.grid(True, axis="y", color=grid,
            linewidth=0.5, linestyle="--", alpha=0.7, zorder=0)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.spines["bottom"].set_visible(True)
    ax.spines["bottom"].set_color(grid)
    ax.tick_params(length=0)


def _embed(fig, parent):
    canvas = FigureCanvasTkAgg(fig, master=parent)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="x", pady=6, padx=4)
    plt.close(fig)


def _section_title(parent, text, t):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.pack(fill="x", padx=6, pady=(14, 2))
    ctk.CTkLabel(frame, text=text,
                 font=("Arial", 12, "bold"),
                 text_color=PALETTE["orange_light"],
                 anchor="w").pack(side="left", padx=6)
    sep_bg = "#2A2A4A" if t["bg"] in ("#1A1A2E", "#0F0F1A") else "#E8D8C8"
    tk.Frame(frame, bg=sep_bg, height=1).pack(
        side="left", fill="x", expand=True, padx=10, pady=10)


def _make_kpi_card(parent, titre, valeur, couleur, t):
    card = ctk.CTkFrame(parent, fg_color=t["card"],
                        corner_radius=14, border_width=1,
                        border_color=couleur)
    card.pack(side="left", padx=7, pady=6, expand=True, fill="both")
    tk.Frame(card, bg=couleur, height=3).pack(fill="x")
    ctk.CTkLabel(card, text=titre, font=("Arial", 10),
                 text_color=t["text"], anchor="w"
                 ).pack(padx=14, pady=(10, 0), anchor="w")
    ctk.CTkLabel(card, text=valeur, font=("Arial", 17, "bold"),
                 text_color=couleur, anchor="w"
                 ).pack(padx=14, pady=(2, 12), anchor="w")


# ══════════════════════════════════════════════════════════
#   AFFICHER HISTORIQUE
# ══════════════════════════════════════════════════════════
def afficher_historique(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t       = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")

    # ── En-tete ──────────────────────────────
    header = ctk.CTkFrame(parent, fg_color=t["card"], corner_radius=0)
    header.pack(fill="x")

    ctk.CTkLabel(header,
                 text="📅  Historique & Statistiques",
                 font=("Arial", 22, "bold"),
                 text_color=PALETTE["orange"]
                 ).pack(side="left", padx=24, pady=14)

    # ── Filtre annee ─────────────────────────
    filter_frame = ctk.CTkFrame(header, fg_color="transparent")
    filter_frame.pack(side="right", padx=20)

    ctk.CTkLabel(filter_frame, text="Annee :",
                 font=("Arial", 12),
                 text_color=t["text"]).pack(side="left", padx=5)

    annees = sorted(set(
        v.date_vente.year for v in session.query(Vente).all()
        if v.date_vente
    ), reverse=True)
    if not annees:
        annees = [date.today().year]

    var_annee = tk.StringVar(value=str(annees[0]))
    cb_annee  = ttk.Combobox(filter_frame, textvariable=var_annee,
                              values=[str(a) for a in annees],
                              width=8, state="readonly")
    cb_annee.pack(side="left", padx=5)

    ctk.CTkButton(
        filter_frame, text="Afficher", width=110,
        fg_color=PALETTE["orange"], hover_color=PALETTE["orange_light"],
        corner_radius=8,
        command=lambda: _charger(
            int(var_annee.get()), cards_frame, graphs_frame, t, is_dark
        )
    ).pack(side="left", padx=10)

    # ── Zone KPI ─────────────────────────────
    cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
    cards_frame.pack(fill="x", padx=16, pady=(10, 0))

    # ── Zone graphiques scrollable ────────────
    graphs_frame = ctk.CTkScrollableFrame(
        parent,
        fg_color=t.get("bg", D_BG),
        scrollbar_button_color=PALETTE["orange"],
        scrollbar_button_hover_color=PALETTE["orange_light"],
    )
    graphs_frame.pack(fill="both", expand=True, padx=16, pady=(4, 12))

    _charger(annees[0], cards_frame, graphs_frame, t, is_dark)


# ══════════════════════════════════════════════════════════
#   CHARGEMENT DONNEES + RENDU
# ══════════════════════════════════════════════════════════
def _charger(annee, cards_frame, graphs_frame, t, is_dark):
    for w in cards_frame.winfo_children():
        w.destroy()
    for w in graphs_frame.winfo_children():
        w.destroy()

    _cfg_mpl(is_dark)

    # ── Données ──────────────────────────────
    ventes = session.query(Vente).filter(
        Vente.date_vente >= date(annee, 1, 1),
        Vente.date_vente <= date(annee, 12, 31)
    ).all()

    ca_total   = round(sum(v.montant_net for v in ventes), 2)
    nb_ventes  = len(ventes)
    marge      = round(ca_total * 0.35, 2)
    panier_moy = round(ca_total / nb_ventes, 2) if nb_ventes else 0

    top_clients = {}
    top_produits = {}
    for v in ventes:
        client  = session.query(Client).filter_by(id=v.client_id).first()
        produit = session.query(Produit).filter_by(id=v.produit_id).first()
        if client:
            top_clients[client.nom] = (top_clients.get(client.nom, 0)
                                       + v.montant_net)
        if produit:
            top_produits[produit.nom] = (top_produits.get(produit.nom, 0)
                                         + v.quantite)

    nb_clients = len(top_clients)

    # ── KPI Cards ────────────────────────────
    kpis = [
        ("CA Total",     f"{ca_total:,.0f} MAD",    PALETTE["orange"]),
        ("Nb Ventes",    str(nb_ventes),             PALETTE["orange_light"]),
        ("Annee",        str(annee),                 PALETTE["teal"]),
        ("Clients actifs", str(nb_clients),          PALETTE["purple"]),
        ("Marge brute",  f"{marge:,.0f} MAD",        PALETTE["green"]),
        ("Panier moyen", f"{panier_moy:,.0f} MAD",   PALETTE["blue"]),
    ]
    for titre, valeur, couleur in kpis:
        _make_kpi_card(cards_frame, titre, valeur, couleur, t)

    # ── G1 : CA mensuel (barres + courbe) ────
    _section_title(graphs_frame, "  Evolution du CA mensuel", t)
    _graph_ca_mensuel(graphs_frame, ventes, annee, is_dark)

    # ── G2 : Top clients ─────────────────────
    _section_title(graphs_frame, "  Top Clients", t)
    _graph_top_clients(graphs_frame, top_clients, is_dark)

    # ── G3 : Top produits ────────────────────
    _section_title(graphs_frame, "  Top Produits vendus", t)
    _graph_top_produits(graphs_frame, top_produits, is_dark)

    # ── G4 : CA cumulatif ────────────────────
    _section_title(graphs_frame, "  CA Cumulatif", t)
    _graph_ca_cumulatif(graphs_frame, ventes, annee, is_dark)

    # ── G5 : Tableau mensuel ─────────────────
    _section_title(graphs_frame, "  Recapitulatif mensuel", t)
    _graph_tableau(graphs_frame, ventes, annee, is_dark, t)


# ══════════════════════════════════════════════════════════
#   GRAPHIQUE 1 — CA MENSUEL
# ══════════════════════════════════════════════════════════
def _graph_ca_mensuel(parent, ventes, annee, is_dark):
    mois_ca = {m: 0.0 for m in range(1, 13)}
    for v in ventes:
        mois_ca[v.date_vente.month] += v.montant_net

    vals  = np.array([mois_ca[m] for m in range(1, 13)], dtype=float)
    x     = np.arange(12)
    max_v = vals.max() if vals.max() > 0 else 1
    bg    = D_BG if is_dark else L_BG
    muted = D_MUTED if is_dark else L_MUTED

    fig, ax = plt.subplots(figsize=(13, 4.0))
    fig.patch.set_facecolor(bg)
    _ax_style(ax, is_dark)

    # Barres dégradées
    cmap   = plt.cm.YlOrRd
    norm   = vals / max_v
    colors = [cmap(0.35 + 0.6 * n) for n in norm]
    bars   = ax.bar(x, vals, color=colors, width=0.62,
                    edgecolor="none", zorder=3)

    # Courbe + area
    gold = PALETTE["orange_gold"]
    ax.fill_between(x, vals, alpha=0.12, color=gold, zorder=2)
    ax.plot(x, vals, color=gold, linewidth=2.2,
            marker="o", markersize=5, zorder=5,
            markerfacecolor=gold, markeredgecolor=bg,
            markeredgewidth=1.5)

    # Étiquettes
    for i, (bar, val) in enumerate(zip(bars, vals)):
        if val > 0:
            ax.text(i, val + max_v * 0.022, _fmt(val),
                    ha="center", va="bottom",
                    fontsize=7.5, color=muted, fontweight="bold")

    # Annotation meilleur mois
    if max_v > 0:
        best = int(vals.argmax())
        ax.annotate("Meilleur mois",
                    xy=(best, vals[best]),
                    xytext=(best, vals[best] + max_v * 0.14),
                    ha="center", fontsize=7.5,
                    color=PALETTE["orange_gold"], fontweight="bold",
                    arrowprops=dict(arrowstyle="-|>",
                                   color=PALETTE["orange_gold"], lw=1.2))

    ax.set_xticks(x)
    ax.set_xticklabels(MOIS)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt))
    ax.set_ylim(0, max_v * 1.30)
    ax.set_ylabel("Chiffre d'affaires (MAD)")
    ax.set_title(f"Evolution du CA mensuel — {annee}",
                 color=D_TEXT if is_dark else L_TEXT,
                 fontsize=12, fontweight="bold",
                 pad=14, loc="left")

    plt.tight_layout(pad=1.8)
    _embed(fig, parent)


# ══════════════════════════════════════════════════════════
#   GRAPHIQUE 2 — TOP CLIENTS
# ══════════════════════════════════════════════════════════
def _graph_top_clients(parent, top_clients, is_dark):
    top5 = sorted(top_clients.items(), key=lambda x: x[1], reverse=True)[:5]

    fig, ax = plt.subplots(figsize=(13, 3.8))
    bg = D_BG if is_dark else L_BG
    fig.patch.set_facecolor(bg)
    _ax_style(ax, is_dark)
    ax.grid(True, axis="x", color=D_GRID if is_dark else L_GRID,
            linewidth=0.5, linestyle="--", alpha=0.7, zorder=0)
    ax.grid(False, axis="y")

    if not top5:
        ax.text(0.5, 0.5, "Aucune donnee", ha="center", va="center",
                fontsize=12, color=D_MUTED if is_dark else L_MUTED)
        ax.set_title("Top 5 Clients", loc="left")
        plt.tight_layout(pad=1.8)
        _embed(fig, parent)
        return

    noms, montants = zip(*top5)
    noms     = list(reversed(noms))
    montants = list(reversed(montants))
    y        = np.arange(len(noms))
    max_m    = max(montants)

    bar_cols = [PALETTE["orange"], PALETTE["orange_light"],
                PALETTE["orange_gold"], PALETTE["teal"], PALETTE["purple"]]
    bar_cols = list(reversed(bar_cols[:len(noms)]))

    bars  = ax.barh(y, montants, color=bar_cols, height=0.52,
                    edgecolor="none", zorder=3)
    text_col = D_TEXT if is_dark else L_TEXT
    total    = sum(montants)

    for bar, val in zip(bars, montants):
        ax.text(val + max_m * 0.015,
                bar.get_y() + bar.get_height() / 2,
                f"{val:,.0f} MAD",
                va="center", fontsize=7.5,
                color=text_col, fontweight="bold")
        ax.text(max_m * 0.015,
                bar.get_y() + bar.get_height() / 2,
                f"{val/total*100:.0f}%",
                va="center", fontsize=7.5,
                color="white", fontweight="bold", zorder=5)

    ax.set_yticks(y)
    ax.set_yticklabels(noms, fontsize=9)
    ax.set_xlim(0, max_m * 1.30)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(_fmt))
    ax.set_title("Top 5 Clients par CA",
                 color=D_TEXT if is_dark else L_TEXT,
                 fontsize=12, fontweight="bold",
                 pad=14, loc="left")

    plt.tight_layout(pad=1.8)
    _embed(fig, parent)


# ══════════════════════════════════════════════════════════
#   GRAPHIQUE 3 — TOP PRODUITS
# ══════════════════════════════════════════════════════════
def _graph_top_produits(parent, top_produits, is_dark):
    top5 = sorted(top_produits.items(), key=lambda x: x[1], reverse=True)[:5]

    fig, ax = plt.subplots(figsize=(13, 3.8))
    bg = D_BG if is_dark else L_BG
    fig.patch.set_facecolor(bg)
    _ax_style(ax, is_dark)

    if not top5:
        ax.text(0.5, 0.5, "Aucune donnee", ha="center", va="center",
                fontsize=12, color=D_MUTED if is_dark else L_MUTED)
        ax.set_title("Top 5 Produits", loc="left")
        plt.tight_layout(pad=1.8)
        _embed(fig, parent)
        return

    noms_p, qtes = zip(*top5)
    x     = np.arange(len(noms_p))
    max_q = max(qtes)
    muted = D_MUTED if is_dark else L_MUTED

    bar_cols = [PALETTE["teal"], PALETTE["orange_light"], PALETTE["orange"],
                PALETTE["orange_gold"], PALETTE["purple"]]

    bars = ax.bar(x, qtes, color=bar_cols[:len(noms_p)],
                  width=0.55, edgecolor="none", zorder=3)

    rank_col = ["#D4AF37", "#A8A9AD", "#CD7F32", "#888888", "#888888"]
    for i, (bar, val) in enumerate(zip(bars, qtes)):
        ax.text(i, val + max_q * 0.025,
                f"{int(val)} unites",
                ha="center", va="bottom",
                fontsize=8, color=muted, fontweight="bold")
        ax.text(i, max_q * 0.04, f"#{i+1}",
                ha="center", va="bottom",
                fontsize=8, color="white", fontweight="bold", zorder=5,
                bbox=dict(boxstyle="round,pad=0.25",
                          facecolor=rank_col[i],
                          edgecolor="none", alpha=0.92))

    ax.set_xticks(x)
    ax.set_xticklabels(noms_p, fontsize=9)
    ax.set_ylim(0, max_q * 1.25)
    ax.set_ylabel("Quantite vendue")
    ax.set_title("Top 5 Produits vendus",
                 color=D_TEXT if is_dark else L_TEXT,
                 fontsize=12, fontweight="bold",
                 pad=14, loc="left")

    plt.tight_layout(pad=1.8)
    _embed(fig, parent)


# ══════════════════════════════════════════════════════════
#   GRAPHIQUE 4 — CA CUMULATIF
# ══════════════════════════════════════════════════════════
def _graph_ca_cumulatif(parent, ventes, annee, is_dark):
    mois_ca = {m: 0.0 for m in range(1, 13)}
    for v in ventes:
        mois_ca[v.date_vente.month] += v.montant_net

    vals  = np.array([mois_ca[m] for m in range(1, 13)], dtype=float)
    cumul = np.cumsum(vals)
    x     = np.arange(12)
    bg    = D_BG if is_dark else L_BG
    muted = D_MUTED if is_dark else L_MUTED

    fig, ax = plt.subplots(figsize=(13, 3.5))
    fig.patch.set_facecolor(bg)
    _ax_style(ax, is_dark)

    color = PALETTE["teal"]
    ax.fill_between(x, cumul, alpha=0.15, color=color, zorder=2)
    ax.plot(x, cumul, color=color, linewidth=2.5, zorder=4)
    ax.scatter(x, cumul, color=color, s=50, zorder=6,
               edgecolors=bg, linewidths=2)

    for i, (xi, yi) in enumerate(zip(x, cumul)):
        if yi > 0:
            offset = 14 if i % 2 == 0 else -22
            va     = "bottom" if i % 2 == 0 else "top"
            ax.annotate(_fmt(yi), xy=(xi, yi),
                        xytext=(0, offset),
                        textcoords="offset points",
                        ha="center", va=va,
                        fontsize=7.5, color=muted, fontweight="bold")

    if cumul[-1] > 0:
        ax.axhline(cumul[-1], color=PALETTE["orange_gold"],
                   linewidth=1, linestyle=":", alpha=0.7)
        ax.text(11.4, cumul[-1] * 1.03,
                f"Total : {_fmt(cumul[-1])} MAD",
                ha="right", fontsize=8,
                color=PALETTE["orange_gold"], fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(MOIS)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt))
    ax.set_ylim(0, max(float(cumul[-1]) * 1.15, 1))
    ax.set_ylabel("CA cumule (MAD)")
    ax.set_title(f"CA cumulatif — {annee}",
                 color=D_TEXT if is_dark else L_TEXT,
                 fontsize=12, fontweight="bold",
                 pad=14, loc="left")

    plt.tight_layout(pad=1.8)
    _embed(fig, parent)


# ══════════════════════════════════════════════════════════
#   GRAPHIQUE 5 — TABLEAU MENSUEL
# ══════════════════════════════════════════════════════════
def _graph_tableau(parent, ventes, annee, is_dark, t):
    mois_ca = {m: 0.0 for m in range(1, 13)}
    mois_nb = {m: 0   for m in range(1, 13)}
    for v in ventes:
        m = v.date_vente.month
        mois_ca[m] += v.montant_net
        mois_nb[m] += 1

    bg    = D_BG    if is_dark else L_BG
    card  = D_CARD  if is_dark else L_CARD
    card2 = D_CARD2 if is_dark else L_CARD2
    grid  = D_GRID  if is_dark else L_GRID
    text  = D_TEXT  if is_dark else L_TEXT

    fig, ax = plt.subplots(figsize=(13, 3.4))
    fig.patch.set_facecolor(bg)
    ax.set_facecolor(bg)
    ax.axis("off")

    total_ca = sum(mois_ca.values())
    col_labels = ["Mois", "CA (MAD)", "Nb Ventes",
                  "Moy / Vente", "Part CA"]
    table_data = []
    for m in range(1, 13):
        ca  = mois_ca[m]
        nb  = mois_nb[m]
        moy = ca / nb if nb > 0 else 0
        pct = ca / total_ca * 100 if total_ca > 0 else 0
        table_data.append([
            f"  {MOIS[m-1]}  ",
            f"  {ca:,.0f}  "   if ca  > 0 else "  -  ",
            f"  {nb}  "        if nb  > 0 else "  -  ",
            f"  {moy:,.0f}  "  if moy > 0 else "  -  ",
            f"  {pct:.1f}%  "  if pct > 0 else "  -  ",
        ])

    tbl = ax.table(
        cellText=table_data,
        colLabels=col_labels,
        cellLoc="center",
        loc="center",
        bbox=[0, 0, 1, 1],
    )
    tbl.auto_set_font_size(False)
    tbl.set_fontsize(9)
    tbl.scale(1, 1.6)

    max_ca = max(mois_ca.values()) if max(mois_ca.values()) > 0 else 1

    for (row, col), cell in tbl.get_celld().items():
        cell.set_edgecolor(grid)
        cell.set_linewidth(0.4)
        if row == 0:
            cell.set_facecolor(PALETTE["orange"])
            cell.set_text_props(color="white",
                                fontweight="bold", fontsize=9)
        elif row % 2 == 0:
            cell.set_facecolor(card2)
            cell.set_text_props(color=text, fontsize=8.5)
        else:
            cell.set_facecolor(card)
            cell.set_text_props(color=text, fontsize=8.5)

        # Dégradé orange sur colonne CA
        if col == 1 and row > 0:
            ca_val = mois_ca[row]
            if ca_val > 0:
                intensity = ca_val / max_ca
                r = int(0xBF + (0xE6 - 0xBF) * intensity)
                g = int(0x36 + (0x51 - 0x36) * intensity)
                b = int(0x0C + (0x00 - 0x0C) * intensity)
                cell.set_facecolor(f"#{r:02x}{g:02x}{b:02x}")
                cell.set_text_props(color="white", fontweight="bold")

    # Résumé total en bas
    nb_total = sum(mois_nb.values())
    panier   = total_ca / nb_total if nb_total > 0 else 0
    ax.text(0.5, -0.05,
            f"Total annuel : {total_ca:,.0f} MAD  |  "
            f"{nb_total} ventes  |  "
            f"Panier moyen : {panier:,.0f} MAD",
            transform=ax.transAxes, ha="center",
            fontsize=8.5, color=PALETTE["orange_light"],
            fontweight="bold")

    ax.set_title(f"Recapitulatif mensuel — {annee}",
                 color=text, fontsize=12, fontweight="bold",
                 pad=10, loc="left", x=0.01)

    plt.tight_layout(pad=1.5)
    _embed(fig, parent)