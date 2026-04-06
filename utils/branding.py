"""
mast.vente — Module de branding centralisé
Gère le logo, les couleurs et l'identité visuelle de toute l'application.
"""
import os
import tkinter as tk
from PIL import Image, ImageTk

APP_NAME      = "mast.vente"
APP_TAGLINE   = "Gestion Ventes & Suivi Commercial"
APP_VERSION   = "2.0"
APP_COPYRIGHT = "© 2025 mast.vente"

BRAND_NAVY    = "#1A2350"
BRAND_ORANGE  = "#F5A623"
BRAND_BLUE    = "#2979C8"
BRAND_LBLUE   = "#5BB8F5"
BRAND_WHITE   = "#FFFFFF"
BRAND_CREAM   = "#FFF8F0"

_BASE     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_LOGO_DIR = os.path.join(_BASE, "assets", "logos")
_cache    = {}

def _logo_path(filename):
    return os.path.join(_LOGO_DIR, filename)

def get_logo_photo(variant="medium"):
    if variant in _cache:
        return _cache[variant]
    names = {
        "login":        "logo_login.png",
        "large":        "logo_large.png",
        "medium":       "logo_medium.png",
        "sidebar_dark": "logo_sidebar_dark.png",
        "sidebar_light":"logo_sidebar_light.png",
        "full":         "logo_full.png",
    }
    try:
        img   = Image.open(_logo_path(names.get(variant, "logo_medium.png")))
        photo = ImageTk.PhotoImage(img)
        _cache[variant] = photo
        return photo
    except Exception as e:
        print(f"[branding] Logo '{variant}' introuvable: {e}")
        return None

def place_logo(parent, variant="medium", bg_color=None, **pack_kwargs):
    try:
        bg = bg_color or parent.cget("bg")
    except Exception:
        bg = BRAND_NAVY
    photo = get_logo_photo(variant)
    if photo:
        lbl = tk.Label(parent, image=photo, bg=bg, bd=0)
        lbl.image = photo
        lbl.pack(**pack_kwargs)
        return lbl
    lbl = tk.Label(parent, text="mast.vente",
                   font=("Arial", 20, "bold"), fg=BRAND_ORANGE, bg=bg)
    lbl.pack(**pack_kwargs)
    return lbl

def get_app_window_title(section=""):
    return f"{APP_NAME} — {section}" if section else APP_NAME
