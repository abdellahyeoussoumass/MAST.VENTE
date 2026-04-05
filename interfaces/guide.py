import customtkinter as ctk
import tkinter as tk
from utils.theme import get_theme


# ══════════════════════════════════════════════════════════
#   CONTENU DU GUIDE — structuré par section
# ══════════════════════════════════════════════════════════
SECTIONS = [
    {
        "id":    "connexion",
        "icone": "🔐",
        "titre": "Connexion & Rôles",
        "couleur": "#E65100",
        "resume": "3 rôles avec accès différents",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Choisir votre rôle",
                "texte":  "Au démarrage, sélectionnez votre rôle dans la liste déroulante.",
                "detail": "• Directeur — accès complet à toutes les sections\n• Commercial — accès sans le Dashboard\n• Admin — accès aux statistiques et à la gestion",
            },
            {
                "numero": "2",
                "titre":  "Saisir le mot de passe",
                "texte":  "Entrez votre mot de passe dans le champ prévu.",
                "detail": "Le mot de passe par défaut est : 1234\nChangez-le dès la première connexion pour sécuriser votre compte.",
            },
            {
                "numero": "3",
                "titre":  "Accéder à l'application",
                "texte":  "Cliquez sur Se connecter pour accéder à votre espace.",
                "detail": "Le menu latéral affiche uniquement les sections autorisées selon votre rôle.",
            },
        ],
        "conseils": [
            "Directeur : idéal pour le suivi global des ventes et finances.",
            "Commercial : focalisé sur les devis, bons de commande et livraisons.",
            "Admin : gestion des données et consultation des statistiques.",
        ],
    },
    {
        "id":    "dashboard",
        "icone": "📊",
        "titre": "Dashboard",
        "couleur": "#FF6D00",
        "resume": "Vue générale des performances",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Choisir une année",
                "texte":  "Utilisez le menu déroulant en haut à droite pour filtrer par année.",
                "detail": "Seules les années ayant des ventes enregistrées apparaissent dans la liste.",
            },
            {
                "numero": "2",
                "titre":  "Lire les cartes KPI",
                "texte":  "6 indicateurs clés s'affichent en haut : CA Total, Total Payé, Total Impayé, Nombre de Ventes, Clients et Marge Brute.",
                "detail": "• CA Total = somme de toutes les ventes nettes de l'année\n• Total Payé = factures avec statut Payée\n• Marge Brute = estimée à 35% du CA Total",
            },
            {
                "numero": "3",
                "titre":  "Analyser les graphiques",
                "texte":  "Faites défiler pour voir les 6 graphiques de l'année sélectionnée.",
                "detail": "• Évolution CA mensuel — barres + courbe de tendance\n• Top 5 Clients — CA par client\n• Payé vs Impayé — donut avec total au centre\n• Top 5 Produits vendus — quantités\n• CA Cumulatif — progression sur l'année\n• Heatmap — activité par semaine et par mois\n• Tableau récapitulatif — CA, ventes et moyenne par mois",
            },
        ],
        "conseils": [
            "Cliquez sur Afficher après avoir changé d'année pour recharger les graphiques.",
            "La heatmap révèle vos semaines les plus actives dans chaque mois.",
            "Le tableau récapitulatif montre la part de CA de chaque mois en pourcentage.",
        ],
    },
    {
        "id":    "produits",
        "icone": "📦",
        "titre": "Produits",
        "couleur": "#2E7D32",
        "resume": "Catalogue de vos produits",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Ajouter un produit",
                "texte":  "Cliquez sur ➕ Ajouter et remplissez le formulaire.",
                "detail": "Champs obligatoires :\n• Référence — code unique du produit\n• Nom — libellé affiché partout\n• Catégorie — pour regrouper et filtrer\n• Prix HT — le prix hors taxe\n• TVA (%) — la taxe sera calculée automatiquement",
            },
            {
                "numero": "2",
                "titre":  "Vérifier le prix TTC",
                "texte":  "Le prix TTC s'affiche en temps réel pendant la saisie du prix HT et de la TVA.",
                "detail": "Formule : Prix TTC = Prix HT × (1 + TVA / 100)\nExemple : 100 MAD HT à 20% TVA = 120 MAD TTC",
            },
            {
                "numero": "3",
                "titre":  "Modifier ou supprimer",
                "texte":  "Sélectionnez un produit dans la liste, puis cliquez sur ✏️ Modifier ou 🗑️ Supprimer.",
                "detail": "Attention : supprimer un produit lié à des ventes ou devis existants peut provoquer des erreurs. Vérifiez d'abord ses dépendances.",
            },
        ],
        "conseils": [
            "Utilisez des catégories claires (ex: Informatique, Mobilier) pour filtrer facilement.",
            "La recherche fonctionne sur le nom et la référence du produit.",
            "Cliquez sur l'en-tête d'une colonne pour trier le tableau.",
        ],
    },
    {
        "id":    "clients",
        "icone": "👥",
        "titre": "Clients",
        "couleur": "#1565C0",
        "resume": "Gestion de votre portefeuille clients",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Ajouter un client manuellement",
                "texte":  "Cliquez sur ➕ Ajouter et renseignez les informations du client.",
                "detail": "Champs disponibles :\n• Nom (obligatoire)\n• Email\n• Téléphone\n• Adresse\n• Ville",
            },
            {
                "numero": "2",
                "titre":  "Importer depuis Excel",
                "texte":  "Cliquez sur 📥 Importer Excel pour charger un fichier .xlsx.",
                "detail": "Format attendu du fichier Excel :\nColonne A : Nom\nColonne B : Email\nColonne C : Téléphone\nColonne D : Adresse\nColonne E : Ville\n\nLa première ligne (en-tête) est ignorée automatiquement.",
            },
            {
                "numero": "3",
                "titre":  "Modifier ou supprimer",
                "texte":  "Sélectionnez un client dans la liste, puis utilisez les boutons Modifier ou Supprimer.",
                "detail": "La modification permet de corriger toutes les informations.\nLa suppression est définitive — vérifiez qu'aucune vente active ne lui est liée.",
            },
        ],
        "conseils": [
            "Préparez votre fichier Excel avec exactement 5 colonnes dans le bon ordre.",
            "La recherche client fonctionne par nom en temps réel.",
            "Double-cliquez sur un client pour voir son détail complet.",
        ],
    },
    {
        "id":    "ventes",
        "icone": "💰",
        "titre": "Ventes",
        "couleur": "#E65100",
        "resume": "Enregistrement des transactions",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Créer une vente",
                "texte":  "Cliquez sur ➕ Ajouter, choisissez le client, le produit et la quantité.",
                "detail": "Le formulaire calcule automatiquement :\n• Prix HT selon le produit sélectionné\n• Montant brut = Prix HT × Quantité\n• Réduction en MAD si un % est saisi\n• Montant net = Brut − Réduction",
            },
            {
                "numero": "2",
                "titre":  "Appliquer une réduction",
                "texte":  "Saisissez un pourcentage de réduction dans le champ prévu.",
                "detail": "Exemple : produit à 1000 MAD, quantité 3, réduction 10%\n• Brut = 3000 MAD\n• Réduction = 300 MAD\n• Net = 2700 MAD",
            },
            {
                "numero": "3",
                "titre":  "Consulter et supprimer",
                "texte":  "Les ventes s'affichent du plus récent au plus ancien. Sélectionnez une ligne pour la supprimer.",
                "detail": "La suppression d'une vente est définitive et n'impacte pas les factures déjà générées.",
            },
        ],
        "conseils": [
            "L'aperçu du montant net se met à jour en temps réel pendant la saisie.",
            "La date est automatiquement fixée à aujourd'hui.",
            "Utilisez la recherche pour retrouver rapidement une vente par client.",
        ],
    },
    {
        "id":    "devis",
        "icone": "📋",
        "titre": "Devis",
        "couleur": "#00838F",
        "resume": "Création et suivi des devis",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Créer un devis",
                "texte":  "Cliquez sur ➕ Ajouter, renseignez le numéro de devis, le client, le produit et la TVA.",
                "detail": "Le formulaire calcule :\n• Prix HT du produit sélectionné\n• Prix TTC avec la TVA saisie\n• Total = Prix TTC × Quantité\n\nL'aperçu s'affiche en temps réel.",
            },
            {
                "numero": "2",
                "titre":  "Choisir le statut",
                "texte":  "Sélectionnez le statut approprié lors de la création ou modification.",
                "detail": "📝 Brouillon — devis en cours de rédaction\n📤 Envoyé — devis transmis au client\n✅ Accepté — devis validé par le client\n❌ Refusé — devis rejeté",
            },
            {
                "numero": "3",
                "titre":  "Envoyer un devis",
                "texte":  "Sélectionnez un devis et cliquez sur 📤 Envoyer.",
                "detail": "La fenêtre d'envoi propose :\n• Email destinataire (pré-rempli si disponible)\n• Message personnalisable\n• Méthode : Email / WhatsApp / Impression\n\nAprès confirmation, le statut passe automatiquement à Envoyé.",
            },
        ],
        "conseils": [
            "Le numéro de devis est libre — utilisez un format cohérent (ex: DEV-2024-001).",
            "Double-cliquez sur un devis pour voir son détail complet.",
            "La recherche fonctionne sur le client, le produit, le numéro et le statut.",
        ],
    },
    {
        "id":    "bons_commande",
        "icone": "🛒",
        "titre": "Bons de Commande",
        "couleur": "#6A1B9A",
        "resume": "Gestion des commandes clients",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Créer un bon de commande",
                "texte":  "Cliquez sur ➕ Ajouter et remplissez : numéro, client, produit, quantité et statut.",
                "detail": "Le total est calculé automatiquement :\nTotal = Prix TTC du produit × Quantité\n\nStatuts disponibles :\n🔄 En cours — commande en traitement\n✅ Payé — commande réglée\n⏳ Pas encore payé — en attente de règlement\n❌ Annulé",
            },
            {
                "numero": "2",
                "titre":  "Importer des bons",
                "texte":  "Cliquez sur 📥 Importer ▾ et choisissez le format de votre fichier.",
                "detail": "Formats supportés :\n• Excel (.xlsx) — colonnes dans l'ordre du tableau\n• CSV (.csv) — séparateur virgule, encodage UTF-8\n• JSON (.json) — tableau sous la clé bons_commande\n\nUne fenêtre de prévisualisation apparaît avant l'import. Seules les lignes avec un client et un produit existants sont importées.",
            },
            {
                "numero": "3",
                "titre":  "Exporter et imprimer",
                "texte":  "Cliquez sur 📤 Exporter ▾ pour sauvegarder, ou 🖨️ Imprimer pour envoyer à l'imprimante.",
                "detail": "Formats d'export disponibles :\n• Excel (.xlsx) — avec mise en forme colorée\n• Word (.docx) — document prêt à envoyer\n• PDF (.pdf) — impression directe\n• CSV et JSON — pour intégration externe\n\nL'impression génère un PDF temporaire et l'ouvre directement.",
            },
        ],
        "conseils": [
            "pip install openpyxl python-docx reportlab pour activer tous les formats.",
            "L'import vérifie les doublons sur le numéro de BC automatiquement.",
            "Double-cliquez sur une ligne pour voir le détail du bon de commande.",
        ],
    },
    {
        "id":    "factures",
        "icone": "🧾",
        "titre": "Factures",
        "couleur": "#C62828",
        "resume": "Facturation et suivi des paiements",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Créer une facture",
                "texte":  "Cliquez sur ➕ Ajouter et sélectionnez client, produit, quantité, TVA et réduction.",
                "detail": "Le calcul automatique donne :\n• Prix HT = prix catalogue\n• Prix TTC = HT × (1 + TVA%)\n• Réduction = Prix TTC × (Remise%)\n• Total final = (Prix TTC − Réduction) × Quantité",
            },
            {
                "numero": "2",
                "titre":  "Choisir le mode de paiement",
                "texte":  "Sélectionnez le mode de paiement et le statut de la facture.",
                "detail": "Modes disponibles :\n💵 Espèces  |  🏦 Virement  |  📄 Chèque  |  💳 Carte\n\nStatuts :\n✅ Payée — règlement reçu\n⏳ En attente — non encore réglée\n❌ Annulée",
            },
            {
                "numero": "3",
                "titre":  "Imprimer / Exporter",
                "texte":  "Sélectionnez une facture et cliquez sur 🖨️ Imprimer pour générer le PDF.",
                "detail": "Le PDF de facture contient :\n• En-tête avec numéro et date\n• Détail client et produit\n• Tableau de calcul avec TVA et réduction\n• Total en bas de page\n• Statut de paiement mis en évidence",
            },
        ],
        "conseils": [
            "Numérotez vos factures de façon séquentielle (ex: FAC-2024-001).",
            "Passez le statut à Payée dès réception du règlement pour un CA exact.",
            "La colonne Total Impayé du Dashboard se base sur les factures non Payées.",
        ],
    },
    {
        "id":    "livraisons",
        "icone": "🚚",
        "titre": "Livraisons",
        "couleur": "#2E7D32",
        "resume": "Bons de livraison et suivi",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Créer un bon de livraison",
                "texte":  "Cliquez sur ➕ Ajouter et renseignez : numéro BL, client, adresse, produit, quantité et statut.",
                "detail": "Le prix TTC est repris automatiquement du produit sélectionné.\nL'aperçu du total (Prix TTC × Quantité) s'affiche en temps réel.",
            },
            {
                "numero": "2",
                "titre":  "Modifier une livraison",
                "texte":  "Sélectionnez une livraison et cliquez sur ✏️ Modifier pour changer toutes les informations.",
                "detail": "Le formulaire de modification est divisé en 3 sections :\n• Informations générales — N° BL, Client, Adresse\n• Produit & Quantité — avec aperçu prix en temps réel\n• Statut — avec indicateur coloré selon l'état\n\nStatuts disponibles :\n⏳ En attente  →  🔄 En cours  →  ✅ Livré  →  ❌ Annulé",
            },
            {
                "numero": "3",
                "titre":  "Exporter et imprimer",
                "texte":  "Utilisez 📥 Importer ▾ et 📤 Exporter ▾ pour les échanges de données, et 🖨️ Imprimer pour l'impression.",
                "detail": "Imports : Excel, CSV, JSON\nExports : Excel, Word, PDF, CSV, JSON\n\nL'impression vous demande : imprimer seulement la sélection ou toutes les livraisons.\nDepuis le détail (double-clic), un bouton Imprimer individuel est disponible.",
            },
        ],
        "conseils": [
            "Mettez à jour le statut à chaque étape pour un suivi précis.",
            "L'adresse de livraison peut différer de l'adresse du client.",
            "Double-cliquez sur une ligne pour voir le détail et modifier ou imprimer directement.",
        ],
    },
    {
        "id":    "historique",
        "icone": "📅",
        "titre": "Historique",
        "couleur": "#FF8F00",
        "resume": "Statistiques des années passées",
        "etapes": [
            {
                "numero": "1",
                "titre":  "Filtrer par année",
                "texte":  "Choisissez une année dans le menu déroulant en haut de la page.",
                "detail": "Seules les années avec des ventes enregistrées sont disponibles dans la liste.",
            },
            {
                "numero": "2",
                "titre":  "Consulter les indicateurs",
                "texte":  "Les statistiques de l'année sélectionnée s'affichent : CA total, nombre de ventes, clients actifs.",
                "detail": "Client actif = client ayant au moins une vente dans l'année sélectionnée.",
            },
            {
                "numero": "3",
                "titre":  "Lire le graphique",
                "texte":  "Le graphique montre l'évolution mensuelle du CA pour comparer les mois.",
                "detail": "Comparez les graphiques d'une année à l'autre pour identifier les tendances saisonnières et les mois les plus performants.",
            },
        ],
        "conseils": [
            "Comparez le CA de différentes années pour mesurer la croissance.",
            "Un mois à 0 MAD signifie qu'aucune vente n'a été enregistrée ce mois-là.",
            "Les données de l'historique ne sont pas modifiables — elles reflètent les ventes réelles.",
        ],
    },
]


# ══════════════════════════════════════════════════════════
#   COMPOSANTS UI
# ══════════════════════════════════════════════════════════
def _badge_numero(parent, numero, couleur, t):
    """Cercle numéroté coloré."""
    badge = tk.Canvas(parent, width=28, height=28,
                      bg=t["bg"], highlightthickness=0)
    badge.create_oval(2, 2, 26, 26, fill=couleur, outline="")
    badge.create_text(14, 14, text=numero,
                      fill="white", font=("Arial", 10, "bold"))
    return badge


def _carte_etape(parent, etape, couleur, t, is_dark):
    """Carte d'une étape avec numéro, titre, texte et détail."""
    card_bg = "#1A1A35" if is_dark else "#FFFFFF"
    border  = couleur

    card = tk.Frame(parent, bg=card_bg,
                    highlightbackground=border,
                    highlightthickness=1)
    card.pack(fill="x", padx=0, pady=4)

    # Ligne haut colorée fine
    tk.Frame(card, bg=couleur, height=2).pack(fill="x")

    inner = tk.Frame(card, bg=card_bg)
    inner.pack(fill="x", padx=14, pady=10)

    # Badge + titre
    top_row = tk.Frame(inner, bg=card_bg)
    top_row.pack(fill="x")

    badge = tk.Canvas(top_row, width=28, height=28,
                      bg=card_bg, highlightthickness=0)
    badge.create_oval(2, 2, 26, 26, fill=couleur, outline="")
    badge.create_text(14, 14, text=etape["numero"],
                      fill="white", font=("Arial", 10, "bold"))
    badge.pack(side="left", padx=(0, 10))

    tk.Label(top_row, text=etape["titre"],
             bg=card_bg, fg=couleur,
             font=("Arial", 12, "bold"),
             anchor="w").pack(side="left", fill="x", expand=True)

    # Texte principal
    fg_text = "#E8E0FF" if is_dark else "#1A1A1A"
    tk.Label(inner, text=etape["texte"],
             bg=card_bg, fg=fg_text,
             font=("Arial", 11),
             justify="left", anchor="w",
             wraplength=580).pack(fill="x", pady=(6, 4))

    # Détail dans un sous-cadre grisé
    if etape.get("detail"):
        detail_bg = "#12122A" if is_dark else "#FFF8F0"
        detail_frame = tk.Frame(inner, bg=detail_bg)
        detail_frame.pack(fill="x", pady=(2, 0))
        fg_detail = "#A0A0C0" if is_dark else "#555555"
        tk.Label(detail_frame, text=etape["detail"],
                 bg=detail_bg, fg=fg_detail,
                 font=("Arial", 10),
                 justify="left", anchor="w",
                 wraplength=560,
                 padx=10, pady=8).pack(fill="x")


def _carte_conseils(parent, conseils, couleur, t, is_dark):
    """Bloc conseils pratiques."""
    tip_bg = "#16213E" if is_dark else "#FFF3E0"
    frame  = tk.Frame(parent, bg=tip_bg,
                      highlightbackground=couleur,
                      highlightthickness=1)
    frame.pack(fill="x", padx=0, pady=(8, 4))

    tk.Frame(frame, bg=couleur, height=2).pack(fill="x")

    header = tk.Frame(frame, bg=tip_bg)
    header.pack(fill="x", padx=12, pady=(8, 4))
    tk.Label(header, text="💡  Conseils pratiques",
             bg=tip_bg, fg=couleur,
             font=("Arial", 11, "bold"),
             anchor="w").pack(side="left")

    fg_tip = "#C0C0D8" if is_dark else "#444444"
    for conseil in conseils:
        row = tk.Frame(frame, bg=tip_bg)
        row.pack(fill="x", padx=12, pady=2)
        tk.Label(row, text="→",
                 bg=tip_bg, fg=couleur,
                 font=("Arial", 10, "bold"),
                 width=2).pack(side="left")
        tk.Label(row, text=conseil,
                 bg=tip_bg, fg=fg_tip,
                 font=("Arial", 10),
                 justify="left", anchor="w",
                 wraplength=560).pack(side="left", fill="x", expand=True)

    tk.Frame(frame, bg=tip_bg, height=6).pack()


# ══════════════════════════════════════════════════════════
#   AFFICHER GUIDE
# ══════════════════════════════════════════════════════════
def afficher_guide(parent):
    for widget in parent.winfo_children():
        widget.destroy()

    t       = get_theme()
    is_dark = t["bg"] in ("#1A1A2E", "#0F0F1A", "#0D0D1A")

    # ── En-tête ──────────────────────────────
    header = ctk.CTkFrame(parent, fg_color=t["card"], corner_radius=0)
    header.pack(fill="x")

    ctk.CTkLabel(
        header,
        text="📖  Guide d'utilisation — VentePro",
        font=("Arial", 22, "bold"),
        text_color="#E65100",
    ).pack(side="left", padx=24, pady=14)

    ctk.CTkLabel(
        header,
        text="Toutes les fonctionnalités expliquées étape par étape",
        font=("Arial", 11),
        text_color="#FF8F00",
    ).pack(side="left", padx=0, pady=14)

    # ── Layout : nav gauche + contenu droit ──
    main_frame = ctk.CTkFrame(parent, fg_color="transparent")
    main_frame.pack(fill="both", expand=True, padx=0, pady=0)

    # ── Barre de navigation latérale ─────────
    nav_bg = "#16213E" if is_dark else "#FFE0B2"
    nav_frame = tk.Frame(main_frame, bg=nav_bg, width=200)
    nav_frame.pack(side="left", fill="y", padx=0, pady=0)
    nav_frame.pack_propagate(False)

    tk.Label(nav_frame, text="  SECTIONS",
             bg=nav_bg,
             fg="#FF8F00",
             font=("Arial", 9, "bold"),
             anchor="w").pack(fill="x", padx=10, pady=(14, 6))

    # ── Zone contenu scrollable ───────────────
    content_outer = ctk.CTkScrollableFrame(
        main_frame,
        fg_color=t["bg"],
        scrollbar_button_color="#E65100",
        scrollbar_button_hover_color="#FF8F00",
    )
    content_outer.pack(side="left", fill="both", expand=True)

    # Référence vers les frames de chaque section
    section_frames = {}

    def _scroll_to(section_id):
        """Scroller jusqu'à la section cliquée."""
        target = section_frames.get(section_id)
        if target:
            target.update_idletasks()
            content_outer._parent_canvas.yview_moveto(0)
            content_outer.update_idletasks()
            # Calcul de la position relative
            y_target = target.winfo_y()
            y_total  = content_outer._parent_canvas.winfo_height()
            canvas_h = content_outer._parent_canvas.bbox("all")
            if canvas_h:
                frac = y_target / canvas_h[3]
                content_outer._parent_canvas.yview_moveto(max(0, frac - 0.02))

    # ── Construire les boutons nav + le contenu ─
    for sec in SECTIONS:
        sec_id  = sec["id"]
        couleur = sec["couleur"]
        titre   = sec["titre"]
        icone   = sec["icone"]

        # Bouton nav
        nav_btn_bg    = nav_bg
        nav_btn_hover = "#252545" if is_dark else "#FFD0A0"
        fg_nav        = "#E8E0FF" if is_dark else "#212121"

        btn_frame = tk.Frame(nav_frame, bg=nav_btn_bg, cursor="hand2")
        btn_frame.pack(fill="x", padx=4, pady=1)

        dot = tk.Frame(btn_frame, bg=couleur, width=4)
        dot.pack(side="left", fill="y")

        btn_lbl = tk.Label(btn_frame,
                           text=f"  {icone}  {titre}",
                           bg=nav_btn_bg, fg=fg_nav,
                           font=("Arial", 10),
                           anchor="w", pady=7)
        btn_lbl.pack(side="left", fill="x", expand=True)

        def _make_click(sid=sec_id, fr=btn_frame, lb=btn_lbl):
            def _enter(e):
                fr.configure(bg=nav_btn_hover)
                lb.configure(bg=nav_btn_hover)
            def _leave(e):
                fr.configure(bg=nav_btn_bg)
                lb.configure(bg=nav_btn_bg)
            def _click(e):
                _scroll_to(sid)
            fr.bind("<Enter>",   _enter)
            fr.bind("<Leave>",   _leave)
            lb.bind("<Enter>",   _enter)
            lb.bind("<Leave>",   _leave)
            fr.bind("<Button-1>", _click)
            lb.bind("<Button-1>", _click)

        _make_click()

        # ── Section dans le contenu ───────────
        sec_frame = tk.Frame(content_outer, bg=t["bg"])
        sec_frame.pack(fill="x", padx=20, pady=(20, 4))
        section_frames[sec_id] = sec_frame

        # En-tête de section
        hdr_sec = tk.Frame(sec_frame, bg=t["bg"])
        hdr_sec.pack(fill="x")

        # Pastille colorée + titre
        tk.Frame(hdr_sec, bg=couleur, width=6).pack(
            side="left", fill="y", padx=(0, 12))

        titre_frame = tk.Frame(hdr_sec, bg=t["bg"])
        titre_frame.pack(side="left", fill="x", expand=True, pady=4)

        fg_titre = "#E8E0FF" if is_dark else "#1A1A1A"
        tk.Label(titre_frame,
                 text=f"{icone}  {titre}",
                 bg=t["bg"], fg=fg_titre,
                 font=("Arial", 17, "bold"),
                 anchor="w").pack(fill="x")

        tk.Label(titre_frame,
                 text=sec["resume"],
                 bg=t["bg"], fg="#FF8F00",
                 font=("Arial", 10, "italic"),
                 anchor="w").pack(fill="x")

        # Séparateur
        sep_bg = "#2A2A4A" if is_dark else "#E8D8C8"
        tk.Frame(sec_frame, bg=sep_bg, height=1).pack(
            fill="x", pady=(8, 12))

        # Étapes
        etapes_label_bg = t["bg"]
        fg_sub = "#A0A0C0" if is_dark else "#888888"
        tk.Label(sec_frame, text="ÉTAPES",
                 bg=t["bg"], fg=fg_sub,
                 font=("Arial", 8, "bold"),
                 anchor="w").pack(fill="x", pady=(0, 4))

        for etape in sec["etapes"]:
            _carte_etape(sec_frame, etape, couleur, t, is_dark)

        # Conseils
        if sec.get("conseils"):
            _carte_conseils(sec_frame, sec["conseils"],
                            couleur, t, is_dark)

        # Séparateur de fin de section
        tk.Frame(content_outer, bg=sep_bg, height=1).pack(
            fill="x", padx=20, pady=8)

    # ── Pied de page ─────────────────────────
    footer_bg = "#16213E" if is_dark else "#FFF0E5"
    footer = tk.Frame(content_outer, bg=footer_bg)
    footer.pack(fill="x", padx=20, pady=(4, 20))

    tk.Label(footer,
             text="VentePro — Guide d'utilisation complet",
             bg=footer_bg, fg="#FF8F00",
             font=("Arial", 10, "bold"),
             pady=10).pack()
    tk.Label(footer,
             text="Pour toute question, contactez votre administrateur.",
             bg=footer_bg,
             fg="#A0A0C0" if is_dark else "#888888",
             font=("Arial", 9),
             pady=2).pack()