import customtkinter as ctk

# ---- Couleurs thème chaleureux ----
THEME = {
    "light": {
        "primary": "#E65100",        # Orange chaud
        "primary_hover": "#BF360C",  # Orange foncé
        "secondary": "#FF8F00",      # Ambre
        "success": "#2E7D32",        # Vert
        "danger": "#C62828",         # Rouge
        "bg": "#FFF8F0",             # Fond crème
        "sidebar": "#E65100",        # Sidebar orange
        "sidebar_hover": "#BF360C",  # Hover sidebar
        "text": "#212121",           # Texte foncé
        "text_light": "#FFFFFF",     # Texte clair
        "card": "#FFE0B2",           # Carte orange clair
    },
    "dark": {
        "primary": "#FF6D00",        # Orange vif
        "primary_hover": "#E65100",  # Orange foncé
        "secondary": "#FFB300",      # Ambre
        "success": "#43A047",        # Vert
        "danger": "#E53935",         # Rouge
        "bg": "#1A1A2E",             # Fond sombre
        "sidebar": "#16213E",        # Sidebar sombre
        "sidebar_hover": "#0F3460",  # Hover sidebar
        "text": "#FFFFFF",           # Texte clair
        "text_light": "#FFFFFF",     # Texte clair
        "card": "#16213E",           # Carte sombre
    }
}

mode_actuel = ["light"]

def get_theme():
    return THEME[mode_actuel[0]]

def basculer_mode():
    mode_actuel[0] = "dark" if mode_actuel[0] == "light" else "light"
    if mode_actuel[0] == "dark":
        ctk.set_appearance_mode("dark")
    else:
        ctk.set_appearance_mode("light")
    return mode_actuel[0]