import smtplib
import json
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ══════════════════════════════════════════════════════════
#   FICHIER DE CONFIGURATION (stocké localement)
# ══════════════════════════════════════════════════════════
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "email_config.json")


def _charger_config():
    """Charge la configuration email depuis le fichier JSON."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "email":      "",
        "mot_de_passe": "",
        "smtp_host":  "smtp.gmail.com",
        "smtp_port":  465,
        "ssl":        True,
    }


def _sauvegarder_config(cfg):
    """Sauvegarde la configuration email."""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def _config_valide(cfg):
    """Vérifie que la config est remplie."""
    return bool(cfg.get("email") and cfg.get("mot_de_passe")
                and "@" in cfg.get("email", ""))


# ══════════════════════════════════════════════════════════
#   FENÊTRE DE CONFIGURATION EMAIL
# ══════════════════════════════════════════════════════════
def ouvrir_config_email(parent=None, callback=None):
    """Ouvre la fenêtre de configuration SMTP."""
    import tkinter as tk
    from tkinter import messagebox, ttk

    try:
        from utils.theme import get_theme
        t = get_theme()
    except Exception:
        t = {
            "bg": "#FFFFFF", "card": "#FFF8F0", "text": "#1A1A1A",
            "text_muted": "#888888", "border": "#E8D8C8",
        }

    cfg = _charger_config()

    win = tk.Toplevel(parent) if parent else tk.Tk()
    win.title("⚙️  Configuration Email")
    win.geometry("520x620")
    win.configure(bg=t["bg"])
    win.resizable(False, False)
    if parent:
        win.grab_set()

    # En-tête
    hdr = tk.Frame(win, bg="#E65100", height=54)
    hdr.pack(fill="x")
    tk.Label(hdr, text="⚙️  Configuration de l'envoi Email",
             bg="#E65100", fg="white",
             font=("Arial", 14, "bold")).pack(pady=14)

    body = tk.Frame(win, bg=t["bg"])
    body.pack(fill="both", expand=True, padx=24, pady=16)

    # ── Guide Gmail ───────────────────────────────────────
    guide_bg = "#1A1A35" if t["bg"] in ("#1A1A2E", "#0F0F1A") else "#FFF3E0"
    guide = tk.Frame(body, bg=guide_bg)
    guide.pack(fill="x", pady=(0, 14))
    tk.Frame(guide, bg="#FF8F00", height=3).pack(fill="x")
    tk.Label(guide,
             text="💡  Pour Gmail, utilisez un Mot de passe d'application :",
             bg=guide_bg, fg="#FF8F00",
             font=("Arial", 10, "bold"),
             anchor="w").pack(fill="x", padx=12, pady=(8, 2))
    for etape in [
        "1. Allez sur myaccount.google.com",
        "2. Sécurité → Vérification en 2 étapes (activer)",
        "3. Cherchez « Mots de passe des applications »",
        "4. Créez un mot de passe pour « Application mail »",
        "5. Copiez le code à 16 caractères généré",
    ]:
        tk.Label(guide, text=f"    {etape}",
                 bg=guide_bg, fg=t["text"],
                 font=("Arial", 9), anchor="w").pack(fill="x", padx=12)
    tk.Frame(guide, bg=guide_bg, height=8).pack()

    # ── Champs de configuration ───────────────────────────
    def _lbl(text):
        tk.Label(body, text=text, bg=t["bg"], fg=t["text"],
                 font=("Arial", 11, "bold"), anchor="w").pack(
                 fill="x", pady=(10, 2))

    def _ent(default="", show=""):
        e = tk.Entry(body, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6, show=show)
        e.pack(fill="x", ipady=5)
        if default:
            e.insert(0, str(default))
        return e

    _lbl("Email expéditeur (votre Gmail)")
    e_email = _ent(cfg.get("email", ""))

    _lbl("Mot de passe d'application (16 caractères)")
    mdp_frame = tk.Frame(body, bg=t["bg"])
    mdp_frame.pack(fill="x")

    e_mdp = tk.Entry(mdp_frame, font=("Arial", 11),
                     bg=t["card"], fg=t["text"],
                     insertbackground=t["text"],
                     relief="flat", bd=6, show="●")
    e_mdp.pack(side="left", fill="x", expand=True, ipady=5)
    if cfg.get("mot_de_passe"):
        e_mdp.insert(0, cfg["mot_de_passe"])

    show_var = tk.BooleanVar(value=False)
    def _toggle_show():
        e_mdp.configure(show="" if show_var.get() else "●")
    tk.Checkbutton(mdp_frame, text="Afficher",
                   variable=show_var,
                   command=_toggle_show,
                   bg=t["bg"], fg=t["text"],
                   selectcolor=t["card"],
                   font=("Arial", 9),
                   activebackground=t["bg"]).pack(side="left", padx=8)

    # ── Options SMTP avancées ─────────────────────────────
    adv_var = tk.BooleanVar(value=False)
    adv_frame = tk.Frame(body, bg=t["bg"])

    def _toggle_adv():
        if adv_var.get():
            adv_frame.pack(fill="x")
        else:
            adv_frame.pack_forget()

    tk.Checkbutton(body, text="⚙  Options SMTP avancées (autre fournisseur)",
                   variable=adv_var,
                   command=_toggle_adv,
                   bg=t["bg"], fg="#FF8F00",
                   selectcolor=t["card"],
                   font=("Arial", 9, "bold"),
                   activebackground=t["bg"]).pack(anchor="w", pady=(10, 0))

    row_smtp = tk.Frame(adv_frame, bg=t["bg"])
    row_smtp.pack(fill="x", pady=(8, 0))

    col1 = tk.Frame(row_smtp, bg=t["bg"])
    col1.pack(side="left", expand=True, fill="x", padx=(0, 8))
    tk.Label(col1, text="Serveur SMTP", bg=t["bg"], fg=t["text"],
             font=("Arial", 10, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_host = tk.Entry(col1, font=("Arial", 10), bg=t["card"], fg=t["text"],
                      insertbackground=t["text"], relief="flat", bd=5)
    e_host.insert(0, cfg.get("smtp_host", "smtp.gmail.com"))
    e_host.pack(fill="x", ipady=4)

    col2 = tk.Frame(row_smtp, bg=t["bg"])
    col2.pack(side="left", width=80)
    tk.Label(col2, text="Port", bg=t["bg"], fg=t["text"],
             font=("Arial", 10, "bold"), anchor="w").pack(fill="x", pady=(0, 2))
    e_port = tk.Entry(col2, font=("Arial", 10), bg=t["card"], fg=t["text"],
                      insertbackground=t["text"], relief="flat", bd=5, width=6)
    e_port.insert(0, str(cfg.get("smtp_port", 465)))
    e_port.pack(fill="x", ipady=4)

    ssl_var = tk.BooleanVar(value=cfg.get("ssl", True))
    tk.Checkbutton(adv_frame, text="Utiliser SSL (port 465)",
                   variable=ssl_var,
                   bg=t["bg"], fg=t["text"],
                   selectcolor=t["card"],
                   font=("Arial", 9),
                   activebackground=t["bg"]).pack(anchor="w", pady=(6, 0))

    # ── Résultat test ─────────────────────────────────────
    result_var = tk.StringVar(value="")
    result_lbl = tk.Label(body, textvariable=result_var,
                          bg=t["bg"], fg="#43A047",
                          font=("Arial", 10, "bold"),
                          wraplength=460, justify="left")
    result_lbl.pack(fill="x", pady=(8, 0))

    # ── Boutons ───────────────────────────────────────────
    sep_col = "#2A2A4A" if t["bg"] in ("#1A1A2E",) else "#E8D8C8"
    tk.Frame(body, bg=sep_col, height=1).pack(fill="x", pady=(12, 0))
    btn_frame = tk.Frame(body, bg=t["bg"])
    btn_frame.pack(fill="x", pady=(10, 4))

    def _tester():
        email = e_email.get().strip()
        mdp   = e_mdp.get().strip()
        if not email or not mdp:
            result_var.set("⚠️  Remplissez l'email et le mot de passe.")
            result_lbl.configure(fg="#FFA000")
            return
        result_var.set("⏳  Test en cours...")
        result_lbl.configure(fg="#1565C0")
        win.update()
        try:
            host = e_host.get().strip() if adv_var.get() else "smtp.gmail.com"
            port = int(e_port.get()) if adv_var.get() else 465
            use_ssl = ssl_var.get() if adv_var.get() else True
            if use_ssl:
                with smtplib.SMTP_SSL(host, port, timeout=10) as s:
                    s.login(email, mdp)
            else:
                with smtplib.SMTP(host, port, timeout=10) as s:
                    s.starttls()
                    s.login(email, mdp)
            result_var.set("✅  Connexion réussie ! Identifiants valides.")
            result_lbl.configure(fg="#43A047")
        except smtplib.SMTPAuthenticationError:
            result_var.set(
                "❌  Authentification échouée.\n"
                "→ Vérifiez que vous utilisez un MOT DE PASSE D'APPLICATION\n"
                "   (pas votre mot de passe Gmail habituel).")
            result_lbl.configure(fg="#E53935")
        except smtplib.SMTPConnectError:
            result_var.set("❌  Impossible de se connecter au serveur SMTP.")
            result_lbl.configure(fg="#E53935")
        except Exception as ex:
            result_var.set(f"❌  Erreur : {ex}")
            result_lbl.configure(fg="#E53935")

    def _sauvegarder():
        email = e_email.get().strip()
        mdp   = e_mdp.get().strip()
        if not email or "@" not in email:
            messagebox.showerror("Erreur", "Email invalide.", parent=win)
            return
        if not mdp:
            messagebox.showerror("Erreur", "Le mot de passe est obligatoire.", parent=win)
            return
        cfg_new = {
            "email":        email,
            "mot_de_passe": mdp,
            "smtp_host":    e_host.get().strip() if adv_var.get() else "smtp.gmail.com",
            "smtp_port":    int(e_port.get()) if adv_var.get() else 465,
            "ssl":          ssl_var.get() if adv_var.get() else True,
        }
        _sauvegarder_config(cfg_new)
        win.destroy()
        messagebox.showinfo("Succes", "Configuration email sauvegardee !")
        if callback:
            callback()

    tk.Button(btn_frame, text="🔌  Tester la connexion",
              command=_tester,
              bg="#FF8F00", fg="white",
              font=("Arial", 11, "bold"),
              relief="flat", cursor="hand2",
              padx=14, pady=8,
              activebackground="#E65100",
              activeforeground="white").pack(side="left", padx=(0, 6))

    tk.Button(btn_frame, text="💾  Sauvegarder",
              command=_sauvegarder,
              bg="#E65100", fg="white",
              font=("Arial", 11, "bold"),
              relief="flat", cursor="hand2",
              padx=14, pady=8,
              activebackground="#BF360C",
              activeforeground="white").pack(side="left", padx=(0, 6))

    tk.Button(btn_frame, text="Annuler",
              command=win.destroy,
              bg=t["card"], fg=t["text"],
              font=("Arial", 11),
              relief="flat", cursor="hand2",
              pady=8).pack(side="left")

    if parent is None:
        win.mainloop()

    return win


# ══════════════════════════════════════════════════════════
#   FONCTION PRINCIPALE — Envoyer la facture par email
# ══════════════════════════════════════════════════════════
def envoyer_facture_email(email_destinataire, nom_client, fichier_pdf, sujet=None, message=None):
    """
    Envoie la facture PDF par email.
    Utilise la configuration sauvegardée dans email_config.json.
    Retourne True si succès, False sinon.
    """
    cfg = _charger_config()

    if not _config_valide(cfg):
        raise ValueError(
            "Email non configure.\n"
            "Cliquez sur Parametres → Configuration Email."
        )

    email_exp = cfg["email"]
    mdp       = cfg["mot_de_passe"]
    host      = cfg.get("smtp_host", "smtp.gmail.com")
    port      = int(cfg.get("smtp_port", 465))
    use_ssl   = cfg.get("ssl", True)

    try:
        msg = MIMEMultipart()
        msg["From"]    = email_exp
        msg["To"]      = email_destinataire
        msg["Subject"] = sujet if sujet else "Votre facture — VentePro"

        # Corps : utiliser le message fourni ou un message par défaut
        if not sujet:
            sujet = "Votre facture — VentePro"
            msg["Subject"] = sujet
        if not message:
            message = (
                f"Bonjour {nom_client},\n\n"
                "Veuillez trouver ci-joint votre facture.\n\n"
                "Cordialement,\n"
                "L'equipe VentePro"
            )
        msg.attach(MIMEText(message, "plain", "utf-8"))

        # Pièce jointe PDF
        if fichier_pdf and os.path.exists(fichier_pdf):
            with open(fichier_pdf, "rb") as f:
                piece = MIMEBase("application", "octet-stream")
                piece.set_payload(f.read())
                encoders.encode_base64(piece)
                piece.add_header(
                    "Content-Disposition",
                    f'attachment; filename="{os.path.basename(fichier_pdf)}"'
                )
                msg.attach(piece)

        # Envoi SMTP
        if use_ssl:
            with smtplib.SMTP_SSL(host, port, timeout=15) as serveur:
                serveur.login(email_exp, mdp)
                serveur.sendmail(email_exp, email_destinataire, msg.as_string())
        else:
            with smtplib.SMTP(host, port, timeout=15) as serveur:
                serveur.ehlo()
                serveur.starttls()
                serveur.login(email_exp, mdp)
                serveur.sendmail(email_exp, email_destinataire, msg.as_string())

        return True

    except smtplib.SMTPAuthenticationError:
        raise ValueError(
            "Authentification Gmail echouee.\n\n"
            "Verifiez que vous utilisez un MOT DE PASSE D'APPLICATION\n"
            "(pas votre mot de passe Gmail habituel).\n\n"
            "Pour le generer :\n"
            "myaccount.google.com → Securite → "
            "Mots de passe des applications"
        )
    except Exception as ex:
        raise ValueError(f"Erreur d'envoi : {ex}")