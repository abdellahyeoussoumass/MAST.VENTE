from sqlalchemy import Column, Integer, String, Float, Date, ForeignKey
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()

# ✅ Table Utilisateurs (Login)
class Utilisateur(Base):
    __tablename__ = "utilisateurs"
    id = Column(Integer, primary_key=True)
    nom = Column(String)
    mot_de_passe = Column(String)
    role = Column(String)  # directeur / commercial / admin

# ✅ Table Clients
class Client(Base):
    __tablename__ = "clients"
    id = Column(Integer, primary_key=True)
    nom = Column(String)
    email = Column(String)
    telephone = Column(String)
    adresse = Column(String)
    ville = Column(String)
    date_creation = Column(Date)

# ✅ Table Produits
class Produit(Base):
    __tablename__ = "produits"
    id = Column(Integer, primary_key=True)
    reference = Column(String)
    nom = Column(String)
    categorie = Column(String)
    prix_ht = Column(Float)
    tva = Column(Float)
    prix_ttc = Column(Float)
    quantite = Column(Integer)

# ✅ Table Ventes
class Vente(Base):
    __tablename__ = "ventes"
    id = Column(Integer, primary_key=True)
    client_id = Column(Integer, ForeignKey("clients.id"))
    produit_id = Column(Integer, ForeignKey("produits.id"))
    date_vente = Column(Date)
    quantite = Column(Integer)
    prix = Column(Float)
    reduction = Column(Float)
    montant_net = Column(Float)
    ville = Column(String)

# ✅ Table Devis
class Devis(Base):
    __tablename__ = "devis"
    id = Column(Integer, primary_key=True)
    numero_devis = Column(String)
    client_id = Column(Integer, ForeignKey("clients.id"))
    produit_id = Column(Integer, ForeignKey("produits.id"))
    categorie = Column(String)
    prix_ht = Column(Float)
    quantite = Column(Integer)
    tva = Column(Float)
    prix_ttc = Column(Float)
    prix_total = Column(Float)
    statut = Column(String)  # Brouillon / Envoyé / Accepté / Refusé
    date_devis = Column(Date)

# ✅ Table Bons de Commande
class BonCommande(Base):
    __tablename__ = "bons_commande"
    id = Column(Integer, primary_key=True)
    numero_bc = Column(String)
    client_id = Column(Integer, ForeignKey("clients.id"))
    produit_id = Column(Integer, ForeignKey("produits.id"))
    categorie = Column(String)
    quantite = Column(Integer)
    prix_ht = Column(Float)
    prix_ttc = Column(Float)
    prix_total = Column(Float)
    statut = Column(String)  # Payé / Pas encore payé
    date_bc = Column(Date)

# ✅ Table Factures
class Facture(Base):
    __tablename__ = "factures"
    id = Column(Integer, primary_key=True)
    numero_facture = Column(String)
    client_id = Column(Integer, ForeignKey("clients.id"))
    prix_ht = Column(Float)
    tva = Column(Float)
    prix_ttc = Column(Float)
    reduction = Column(Float)
    statut = Column(String)
    mode_paiement = Column(String)

# ✅ Table Livraisons
class Livraison(Base):
    __tablename__ = "livraisons"
    id = Column(Integer, primary_key=True)
    numero_bl = Column(String)
    client_id = Column(Integer, ForeignKey("clients.id"))
    devis_id = Column(Integer, ForeignKey("devis.id"))
    adresse = Column(String)
    produit_id = Column(Integer, ForeignKey("produits.id"))
    prix = Column(Float)
    quantite = Column(Integer)
    statut = Column(String)