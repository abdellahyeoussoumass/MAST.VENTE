from fpdf import FPDF
from database.db import session
from database.models import Vente, Client, Produit, Facture
from datetime import date, datetime
import os

class RapportPDF(FPDF):
    def header(self):
        self.set_fill_color(230, 81, 0)
        self.rect(0, 0, 210, 30, 'F')
        self.set_font("Arial", "B", 20)
        self.set_text_color(255, 255, 255)
        self.cell(0, 15, "VentePro — Rapport Mensuel", ln=True, align="C")
        self.set_font("Arial", size=11)
        self.cell(0, 10,
                  f"Généré le : {datetime.now().strftime('%d/%m/%Y à %H:%M')}",
                  ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 9)
        self.set_text_color(150, 150, 150)
        self.cell(0, 10,
                  f"© VentePro — Page {self.page_no()}",
                  align="C")

    def titre_section(self, titre):
        self.set_fill_color(255, 143, 0)
        self.set_text_color(255, 255, 255)
        self.set_font("Arial", "B", 13)
        self.cell(0, 10, titre, ln=True, fill=True)
        self.set_text_color(0, 0, 0)
        self.ln(3)

    def ligne_tableau(self, colonnes, largeurs, header=False):
        if header:
            self.set_fill_color(230, 81, 0)
            self.set_text_color(255, 255, 255)
            self.set_font("Arial", "B", 10)
        else:
            self.set_fill_color(255, 240, 220)
            self.set_text_color(0, 0, 0)
            self.set_font("Arial", size=9)

        for col, larg in zip(colonnes, largeurs):
            self.cell(larg, 8, str(col), border=1,
                      fill=True, align="C")
        self.ln()
        self.set_text_color(0, 0, 0)

def generer_rapport(mois, annee):
    pdf = RapportPDF()
    pdf.add_page()

    # ---- Données ----
    ventes = session.query(Vente).filter(
        Vente.date_vente >= date(annee, mois, 1),
        Vente.date_vente <= date(
            annee, mois,
            __import__('calendar').monthrange(annee, mois)[1]
        )
    ).all()

    factures = session.query(Facture).all()
    clients = session.query(Client).all()
    produits = session.query(Produit).all()

    ca_total = round(sum(v.montant_net for v in ventes), 2)
    total_paye = round(sum(
        f.prix_ttc for f in factures if f.statut == "Payée"), 2)
    total_impaye = round(sum(
        f.prix_ttc for f in factures if f.statut != "Payée"), 2)
    marge_brute = round(ca_total * 0.35, 2)
    nb_ventes = len(ventes)

    mois_noms = ["", "Janvier", "Février", "Mars", "Avril", "Mai",
                 "Juin", "Juillet", "Août", "Septembre",
                 "Octobre", "Novembre", "Décembre"]

    # ---- Titre période ----
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(230, 81, 0)
    pdf.cell(0, 10,
             f"Période : {mois_noms[mois]} {annee}",
             ln=True, align="C")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(5)

    # ---- Section KPI ----
    pdf.titre_section("📊 Indicateurs Clés de Performance")

    kpis = [
        ("CA Total", f"{ca_total} MAD"),
        ("Total Payé", f"{total_paye} MAD"),
        ("Total Impayé", f"{total_impaye} MAD"),
        ("Marge Brute", f"{marge_brute} MAD"),
        ("Nombre de Ventes", str(nb_ventes)),
        ("Nombre de Clients", str(len(clients))),
        ("Nombre de Produits", str(len(produits))),
    ]

    pdf.set_font("Arial", size=11)
    for label, valeur in kpis:
        pdf.set_fill_color(255, 240, 220)
        pdf.cell(100, 9, f"  {label}", border=1, fill=True)
        pdf.set_fill_color(230, 81, 0)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(90, 9, valeur, border=1, fill=True, align="C")
        pdf.set_text_color(0, 0, 0)
        pdf.ln()

    pdf.ln(8)

    # ---- Section Ventes ----
    pdf.titre_section("💰 Détail des Ventes du Mois")

    if ventes:
        colonnes = ["ID", "Client", "Produit", "Quantité", "Prix", "Réduction", "Montant Net", "Ville"]
        largeurs = [12, 35, 35, 18, 20, 20, 28, 22]
        pdf.ligne_tableau(colonnes, largeurs, header=True)

        for v in ventes:
            client = session.query(Client).filter_by(id=v.client_id).first()
            produit = session.query(Produit).filter_by(id=v.produit_id).first()
            pdf.ligne_tableau([
                v.id,
                client.nom[:15] if client else "N/A",
                produit.nom[:15] if produit else "N/A",
                v.quantite,
                f"{v.prix} MAD",
                f"{v.reduction}%",
                f"{v.montant_net} MAD",
                v.ville or "N/A"
            ], largeurs)
    else:
        pdf.set_font("Arial", "I", 11)
        pdf.cell(0, 10, "Aucune vente ce mois.", ln=True)

    pdf.ln(8)

    # ---- Top 5 Clients ----
    pdf.titre_section("🏆 Top 5 Clients")

    top_clients = {}
    for v in ventes:
        client = session.query(Client).filter_by(id=v.client_id).first()
        if client:
            top_clients[client.nom] = top_clients.get(
                client.nom, 0) + v.montant_net

    top5 = sorted(top_clients.items(),
                  key=lambda x: x[1], reverse=True)[:5]

    if top5:
        colonnes = ["Rang", "Client", "CA Total (MAD)"]
        largeurs = [20, 100, 70]
        pdf.ligne_tableau(colonnes, largeurs, header=True)
        for rang, (nom, ca) in enumerate(top5, 1):
            pdf.ligne_tableau([f"#{rang}", nom, f"{round(ca, 2)} MAD"],
                              largeurs)
    else:
        pdf.set_font("Arial", "I", 11)
        pdf.cell(0, 10, "Aucune donnée disponible.", ln=True)

    pdf.ln(8)

    # ---- Top 5 Produits ----
    pdf.titre_section("📦 Top 5 Produits Vendus")

    top_produits = {}
    for v in ventes:
        produit = session.query(Produit).filter_by(id=v.produit_id).first()
        if produit:
            top_produits[produit.nom] = top_produits.get(
                produit.nom, 0) + v.quantite

    top5p = sorted(top_produits.items(),
                   key=lambda x: x[1], reverse=True)[:5]

    if top5p:
        colonnes = ["Rang", "Produit", "Quantité Vendue"]
        largeurs = [20, 100, 70]
        pdf.ligne_tableau(colonnes, largeurs, header=True)
        for rang, (nom, qte) in enumerate(top5p, 1):
            pdf.ligne_tableau([f"#{rang}", nom, str(qte)], largeurs)
    else:
        pdf.set_font("Arial", "I", 11)
        pdf.cell(0, 10, "Aucune donnée disponible.", ln=True)

    # ---- Sauvegarder ----
    os.makedirs("rapports", exist_ok=True)
    nom_fichier = f"rapports/rapport_{mois_noms[mois]}_{annee}.pdf"
    pdf.output(nom_fichier)
    return nom_fichier