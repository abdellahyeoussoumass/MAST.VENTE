import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
from database.db import session
from database.models import Vente, Client, Produit, Facture
from datetime import date
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.gridspec import GridSpec
import numpy as np
from utils.theme import get_theme


# ══════════════════════════════════════════════════════════
#   PALETTE & CONSTANTES
# ══════════════════════════════════════════════════════════
PALETTE = {
    "orange":       "#E65100",
    "orange_light": "#FF8F00",
    "orange_gold":  "#FFB300",
    "green":        "#2E7D32",
    "green_light":  "#43A047",
    "red":          "#C62828",
    "red_light":    "#E53935",
    "purple":       "#6A1B9A",
    "teal":         "#00838F",
    "blue":         "#1565C0",
}

D_BG     = "#0D0D1A"
D_CARD   = "#12122A"
D_CARD2  = "#1A1A35"
D_GRID   = "#22223A"
D_TEXT   = "#E8E0FF"
D_MUTED  = "#7070A0"

L_BG     = "#FFF8F0"
L_CARD   = "#FFFFFF"
L_CARD2  = "#FFF0E5"
L_GRID   = "#EEE5D8"
L_TEXT   = "#1A1A1A"
L_MUTED  = "#777777"

DARK_BG    = D_BG
CARD_BG    = D_CARD
CARD_BG2   = D_CARD2
ACCENT     = "#E65100"
TEXT_MAIN  = D_TEXT
TEXT_MUTED = D_MUTED
GRID_COLOR = D_GRID

MOIS_LABELS = ["Jan", "Fév", "Mar", "Avr", "Mai", "Jun",
               "Jul", "Aoû", "Sep", "Oct", "Nov", "Déc"]


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
        "axes.titlepad":      14,
        "axes.titlesize":     12,
        "axes.titleweight":   "bold",
        "axes.labelsize":     9,
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
        "xtick.major.pad":    6,
        "ytick.major.pad":    6,
        "text.color":         text,
        "font.family":        "DejaVu Sans",
        "font.size":          9,
        "axes.spines.top":    False,
        "axes.spines.right":  False,
        "axes.spines.left":   False,
        "axes.spines.bottom": True,
        "figure.dpi":         110,
        "legend.frameon":     False,
        "legend.fontsize":    8.5,
    })


# ══════════════════════════════════════════════════════════
#   HELPERS
# ══════════════════════════════════════════════════════════
def _embed(fig, parent):
    canvas = FigureCanvasTkAgg(fig, master=parent)
    canvas.draw()
    canvas.get_tk_widget().pack(fill="x", pady=6, padx=2)
    plt.close(fig)


def _fmt_mad(val, _=None):
    if val >= 1_000_000:
        return f"{val/1_000_000:.1f}M"
    if val >= 1_000:
        return f"{val/1_000:.0f}k"
    return f"{int(val)}"


def _ax_style(ax, is_dark):
    bg   = D_CARD if is_dark else L_CARD
    grid = D_GRID if is_dark else L_GRID
    ax.set_facecolor(bg)
    ax.grid(True, axis="y", color=grid, linewidth=0.5,
            linestyle="--", alpha=0.7, zorder=0)
    ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.spines["bottom"].set_visible(True)
    ax.spines["bottom"].set_color(grid)
    ax.tick_params(length=0)


def _title(ax, text, is_dark):
    col = D_TEXT if is_dark else L_TEXT
    ax.set_title(text, color=col, fontsize=12,
                 fontweight="bold", pad=14, loc="left")


def _section_title(parent, text, t):
    frame = ctk.CTkFrame(parent, fg_color="transparent")
    frame.pack(fill="x", padx=6, pady=(18, 2))
    ctk.CTkLabel(
        frame, text=text,
        font=("Arial", 12, "bold"),
        text_color=PALETTE["orange_light"],
        anchor="w",
    ).pack(side="left", padx=6)
    sep = tk.Frame(frame, bg=PALETTE["orange_light"], height=1)
    sep.pack(side="left", fill="x", expand=True, padx=10, pady=10)


# ══════════════════════════════════════════════════════════
#   CARTES KPI
# ══════════════════════════════════════════════════════════
def _make_kpi_card(parent, titre, valeur, couleur, t):
    card = ctk.CTkFrame(
        parent, fg_color=t["card"],
        corner_radius=14, border_width=1,
        border_color=couleur,
    )
    card.pack(side="left", padx=7, pady=6, expand=True, fill="both")
    tk.Frame(card, bg=couleur, height=3).pack(fill="x")
    ctk.CTkLabel(card, text=titre, font=("Arial", 10),
                 text_color=t["text"], anchor="w"
                 ).pack(padx=14, pady=(10, 0), anchor="w")
    ctk.CTkLabel(card, text=valeur, font=("Arial", 17, "bold"),
                 text_color=couleur, anchor="w"
                 ).pack(padx=14, pady=(2, 12), anchor="w")


# ══════════════════════════════════════════════════════════
#   G1 — CA MENSUEL  barres + courbe area
# ══════════════════════════════════════════════════════════
def _graph_ca_mensuel(graphs_frame, ventes, annee, is_dark):
    mois_ca = {m: 0.0 for m in range(1, 13)}
    for v in ventes:
        mois_ca[v.date_vente.month] += v.montant_net

    vals  = np.array([mois_ca[m] for m in range(1, 13)], dtype=float)
    x     = np.arange(12)
    max_v = vals.max() if vals.max() > 0 else 1

    fig, ax = plt.subplots(figsize=(13, 4.2))
    fig.patch.set_facecolor(D_BG if is_dark else L_BG)
    _ax_style(ax, is_dark)

    # Barres colorées par intensité
    norm   = vals / max_v
    cmap   = plt.cm.YlOrRd
    colors = [cmap(0.35 + 0.6 * n) for n in norm]
    bars   = ax.bar(x, vals, color=colors, width=0.62,
                    edgecolor="none", zorder=3)

    # Area + courbe
    gold = PALETTE["orange_gold"]
    bg   = D_BG if is_dark else L_BG
    ax.fill_between(x, vals, alpha=0.12, color=gold, zorder=2)
    ax.plot(x, vals, color=gold, linewidth=2.2, marker="o",
            markersize=5, zorder=5,
            markerfacecolor=gold, markeredgecolor=bg,
            markeredgewidth=1.5)

    muted = D_MUTED if is_dark else L_MUTED
    for i, (bar, val) in enumerate(zip(bars, vals)):
        if val > 0:
            ax.text(i, val + max_v * 0.022, _fmt_mad(val),
                    ha="center", va="bottom",
                    fontsize=7.5, color=muted, fontweight="bold")

    # Annotation meilleur mois
    if max_v > 0:
        best = int(vals.argmax())
        ax.annotate(
            "Meilleur mois",
            xy=(best, vals[best]),
            xytext=(best, vals[best] + max_v * 0.14),
            ha="center", fontsize=7.5,
            color=PALETTE["orange_gold"], fontweight="bold",
            arrowprops=dict(arrowstyle="-|>",
                            color=PALETTE["orange_gold"], lw=1.2),
        )

    ax.set_xticks(x)
    ax.set_xticklabels(MOIS_LABELS)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_mad))
    ax.set_ylim(0, max_v * 1.28)
    ax.set_ylabel("Chiffre d'affaires (MAD)")
    _title(ax, f"Evolution du CA mensuel — {annee}", is_dark)
    plt.tight_layout(pad=1.8)
    _embed(fig, graphs_frame)


# ══════════════════════════════════════════════════════════
#   G2 — TOP 5 CLIENTS  barres horizontales
# ══════════════════════════════════════════════════════════
def _graph_top_clients(ax, ventes, is_dark):
    top_clients = {}
    for v in ventes:
        client = session.query(Client).filter_by(id=v.client_id).first()
        if client:
            top_clients[client.nom] = (top_clients.get(client.nom, 0)
                                       + v.montant_net)

    top5 = sorted(top_clients.items(), key=lambda x: x[1], reverse=True)[:5]
    _ax_style(ax, is_dark)
    ax.grid(True, axis="x",
            color=D_GRID if is_dark else L_GRID,
            linewidth=0.5, linestyle="--", alpha=0.7, zorder=0)
    ax.grid(False, axis="y")

    if not top5:
        ax.text(0.5, 0.5, "Aucune donnée", ha="center", va="center",
                fontsize=11, color=D_MUTED if is_dark else L_MUTED)
        _title(ax, "Top 5 Clients", is_dark)
        return

    noms, montants   = zip(*top5)
    noms     = list(reversed(noms))
    montants = list(reversed(montants))
    y        = np.arange(len(noms))
    max_m    = max(montants)

    bar_cols = [PALETTE["orange"], PALETTE["orange_light"],
                PALETTE["orange_gold"], PALETTE["teal"], PALETTE["purple"]]
    bar_cols = list(reversed(bar_cols[:len(noms)]))

    bars = ax.barh(y, montants, color=bar_cols,
                   height=0.52, edgecolor="none", zorder=3)

    text_col = D_TEXT if is_dark else L_TEXT
    total    = sum(montants)
    for bar, val in zip(bars, montants):
        # Valeur fin de barre
        ax.text(val + max_m * 0.015,
                bar.get_y() + bar.get_height() / 2,
                f"{val:,.0f} MAD",
                va="center", fontsize=7.5,
                color=text_col, fontweight="bold")
        # Pourcentage dans la barre
        pct = val / total * 100
        ax.text(max_m * 0.015,
                bar.get_y() + bar.get_height() / 2,
                f"{pct:.0f}%",
                va="center", fontsize=7.5,
                color="white", fontweight="bold", zorder=5)

    ax.set_yticks(y)
    ax.set_yticklabels(noms, fontsize=9)
    ax.set_xlim(0, max_m * 1.30)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(_fmt_mad))
    _title(ax, "Top 5 Clients", is_dark)


# ══════════════════════════════════════════════════════════
#   G3 — PAYÉ vs IMPAYÉ  donut
# ══════════════════════════════════════════════════════════
def _graph_paye_impaye(ax, total_paye, total_impaye, is_dark):
    bg = D_CARD if is_dark else L_CARD
    ax.set_facecolor(bg)

    if total_paye == 0 and total_impaye == 0:
        ax.text(0.5, 0.5, "Aucune donnée", ha="center", va="center",
                fontsize=11, color=D_MUTED if is_dark else L_MUTED)
        _title(ax, "Paye vs Impaye", is_dark)
        return

    sizes  = [total_paye, total_impaye]
    colors = [PALETTE["green_light"], PALETTE["red_light"]]

    wedges, _, autotexts = ax.pie(
        sizes, colors=colors, autopct="%1.1f%%",
        startangle=90, pctdistance=0.78,
        wedgeprops={"width": 0.52, "edgecolor": bg, "linewidth": 3},
    )

    text_col = D_TEXT if is_dark else L_TEXT
    for at in autotexts:
        at.set_fontsize(9.5)
        at.set_color(text_col)
        at.set_fontweight("bold")

    total = total_paye + total_impaye
    ax.text(0, 0.1, _fmt_mad(total), ha="center", va="center",
            fontsize=13, fontweight="bold",
            color=PALETTE["orange_light"])
    ax.text(0, -0.15, "MAD Total", ha="center", va="center",
            fontsize=8, color=D_MUTED if is_dark else L_MUTED)

    patches = [
        mpatches.Patch(color=PALETTE["green_light"],
                       label=f"Paye  {total_paye:,.0f} MAD"),
        mpatches.Patch(color=PALETTE["red_light"],
                       label=f"Impaye  {total_impaye:,.0f} MAD"),
    ]
    ax.legend(handles=patches, loc="lower center",
              bbox_to_anchor=(0.5, -0.14), ncol=1,
              fontsize=8.5, frameon=False,
              labelcolor=D_MUTED if is_dark else L_MUTED)
    _title(ax, "Paye vs Impaye", is_dark)


# ══════════════════════════════════════════════════════════
#   G4 — TOP PRODUITS  barres + badges rang
# ══════════════════════════════════════════════════════════
def _graph_top_produits(graphs_frame, ventes, is_dark):
    top_produits = {}
    for v in ventes:
        produit = session.query(Produit).filter_by(id=v.produit_id).first()
        if produit:
            top_produits[produit.nom] = (top_produits.get(produit.nom, 0)
                                         + v.quantite)

    fig, ax = plt.subplots(figsize=(13, 3.8))
    fig.patch.set_facecolor(D_BG if is_dark else L_BG)
    _ax_style(ax, is_dark)

    if not top_produits:
        ax.text(0.5, 0.5, "Aucune donnee", ha="center", va="center",
                fontsize=12, color=D_MUTED if is_dark else L_MUTED)
        _title(ax, "Top 5 Produits vendus", is_dark)
        plt.tight_layout(pad=1.8)
        _embed(fig, graphs_frame)
        return

    top5p = sorted(top_produits.items(), key=lambda x: x[1], reverse=True)[:5]
    noms_p, qtes = zip(*top5p)
    x     = np.arange(len(noms_p))
    max_q = max(qtes)

    bar_cols = [PALETTE["teal"], PALETTE["orange_light"], PALETTE["orange"],
                PALETTE["orange_gold"], PALETTE["purple"]]

    bars = ax.bar(x, qtes, color=bar_cols[:len(noms_p)],
                  width=0.55, edgecolor="none", zorder=3)

    muted    = D_MUTED if is_dark else L_MUTED
    rank_col = ["#D4AF37", "#A8A9AD", "#CD7F32", "#888888", "#888888"]

    for i, (bar, val) in enumerate(zip(bars, qtes)):
        # Label valeur
        ax.text(i, val + max_q * 0.025,
                f"{int(val)} unites",
                ha="center", va="bottom",
                fontsize=8, color=muted, fontweight="bold")
        # Badge rang
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
    _title(ax, "Top 5 Produits vendus", is_dark)
    plt.tight_layout(pad=1.8)
    _embed(fig, graphs_frame)


# ══════════════════════════════════════════════════════════
#   G5 — CA CUMULATIF  area chart
# ══════════════════════════════════════════════════════════
def _graph_ca_cumulatif(graphs_frame, ventes, annee, is_dark):
    mois_ca = {m: 0.0 for m in range(1, 13)}
    for v in ventes:
        mois_ca[v.date_vente.month] += v.montant_net

    vals  = np.array([mois_ca[m] for m in range(1, 13)], dtype=float)
    cumul = np.cumsum(vals)
    x     = np.arange(12)

    fig, ax = plt.subplots(figsize=(13, 3.5))
    fig.patch.set_facecolor(D_BG if is_dark else L_BG)
    _ax_style(ax, is_dark)

    color = PALETTE["teal"]
    bg    = D_BG if is_dark else L_BG
    muted = D_MUTED if is_dark else L_MUTED

    ax.fill_between(x, cumul, alpha=0.15, color=color, zorder=2)
    ax.plot(x, cumul, color=color, linewidth=2.5, zorder=4)
    ax.scatter(x, cumul, color=color, s=50, zorder=6,
               edgecolors=bg, linewidths=2)

    for i, (xi, yi) in enumerate(zip(x, cumul)):
        if yi > 0:
            offset = 14 if i % 2 == 0 else -22
            va     = "bottom" if i % 2 == 0 else "top"
            ax.annotate(_fmt_mad(yi), xy=(xi, yi),
                        xytext=(0, offset),
                        textcoords="offset points",
                        ha="center", va=va,
                        fontsize=7.5, color=muted, fontweight="bold")

    if cumul[-1] > 0:
        ax.axhline(cumul[-1], color=PALETTE["orange_gold"],
                   linewidth=1, linestyle=":", alpha=0.7)
        ax.text(11.4, cumul[-1] * 1.03,
                f"Total : {_fmt_mad(cumul[-1])} MAD",
                ha="right", fontsize=8,
                color=PALETTE["orange_gold"], fontweight="bold")

    ax.set_xticks(x)
    ax.set_xticklabels(MOIS_LABELS)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(_fmt_mad))
    ax.set_ylim(0, max(float(cumul[-1]) * 1.15, 1))
    ax.set_ylabel("CA cumule (MAD)")
    _title(ax, f"CA cumulatif — {annee}", is_dark)
    plt.tight_layout(pad=1.8)
    _embed(fig, graphs_frame)




# ══════════════════════════════════════════════════════════
#   FONCTION PRINCIPALE
# ══════════════════════════════════════════════════════════
def afficher_dashboard(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t = get_theme()

    header = ctk.CTkFrame(parent, fg_color=t["card"], corner_radius=0)
    header.pack(fill="x")

    ctk.CTkLabel(
        header,
        text="📊  Dashboard  —  Vue Generale",
        font=("Arial", 22, "bold"),
        text_color=PALETTE["orange"],
    ).pack(side="left", padx=24, pady=14)

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
        command=lambda: charger_dashboard(
            int(var_annee.get()), cards_frame, graphs_frame
        )
    ).pack(side="left", padx=10)

    cards_frame = ctk.CTkFrame(parent, fg_color="transparent")
    cards_frame.pack(fill="x", padx=16, pady=(10, 0))

    graphs_frame = ctk.CTkScrollableFrame(
        parent,
        fg_color=t.get("bg", D_BG),
        scrollbar_button_color=PALETTE["orange"],
        scrollbar_button_hover_color=PALETTE["orange_light"],
    )
    graphs_frame.pack(fill="both", expand=True, padx=16, pady=(4, 12))

    charger_dashboard(annees[0], cards_frame, graphs_frame)


# ══════════════════════════════════════════════════════════
#   CHARGEMENT DONNEES + RENDU
# ══════════════════════════════════════════════════════════
def charger_dashboard(annee, cards_frame, graphs_frame):
    for widget in cards_frame.winfo_children():
        widget.destroy()
    for widget in graphs_frame.winfo_children():
        widget.destroy()

    t       = get_theme()
    is_dark = t.get("bg", D_BG) in (D_BG, "#1A1A2E", "#0F0F1A")

    _cfg_mpl(is_dark)

    # ── Donnees ─────────────────────────────
    ventes = session.query(Vente).filter(
        Vente.date_vente >= date(annee, 1, 1),
        Vente.date_vente <= date(annee, 12, 31)
    ).all()

    factures = session.query(Facture).all()
    produits = session.query(Produit).all()

    ca_total     = round(sum(v.montant_net for v in ventes), 2)
    nb_ventes    = len(ventes)
    total_paye   = round(sum(f.prix_ttc for f in factures
                             if f.statut == "Payee"), 2)
    total_impaye = round(sum(f.prix_ttc for f in factures
                             if f.statut != "Payee"), 2)
    nb_clients   = session.query(Client).count()
    marge_brute  = round(ca_total * 0.35, 2)

    # ── KPI ─────────────────────────────────
    kpis = [
        ("CA Total",     f"{ca_total:,.0f} MAD",     PALETTE["orange"]),
        ("Total Paye",   f"{total_paye:,.0f} MAD",   PALETTE["green"]),
        ("Total Impaye", f"{total_impaye:,.0f} MAD",  PALETTE["red"]),
        ("Ventes",       str(nb_ventes),              PALETTE["orange_light"]),
        ("Clients",      str(nb_clients),             PALETTE["purple"]),
        ("Marge Brute",  f"{marge_brute:,.0f} MAD",  PALETTE["teal"]),
    ]
    for titre, valeur, couleur in kpis:
        _make_kpi_card(cards_frame, titre, valeur, couleur, t)

    # ── Graphiques ──────────────────────────
    _section_title(graphs_frame, "  Analyse des Ventes", t)
    _graph_ca_mensuel(graphs_frame, ventes, annee, is_dark)

    _section_title(graphs_frame, "  Clients & Paiements", t)
    bg  = D_BG if is_dark else L_BG
    fig2 = plt.figure(figsize=(13, 4.5))
    fig2.patch.set_facecolor(bg)
    gs  = GridSpec(1, 2, figure=fig2, wspace=0.35)
    ax2 = fig2.add_subplot(gs[0, 0])
    ax3 = fig2.add_subplot(gs[0, 1])
    _graph_top_clients(ax2, ventes, is_dark)
    _graph_paye_impaye(ax3, total_paye, total_impaye, is_dark)
    plt.tight_layout(pad=2.0)
    _embed(fig2, graphs_frame)

    _section_title(graphs_frame, "  Produits", t)
    _graph_top_produits(graphs_frame, ventes, is_dark)

    _section_title(graphs_frame, "  CA Cumulatif", t)
    _graph_ca_cumulatif(graphs_frame, ventes, annee, is_dark)