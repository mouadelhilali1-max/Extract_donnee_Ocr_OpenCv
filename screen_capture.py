import os
import time

import pygetwindow as gw
import mss
import mss.tools


# =====================================================
# CATIA
# =====================================================

def get_catia_window():
    windows = gw.getWindowsWithTitle("CATIA")

    if len(windows) == 0:
        return None

    return windows[0]


def activate_catia_window():

    window = get_catia_window()

    if window is None:
        print("CATIA non trouvée")
        return None

    try:

        if window.isMinimized:
            window.restore()

        try:
            window.activate()
        except:
            pass

        time.sleep(0.10)

        return window

    except Exception as e:

        print("Erreur :", e)
        return None


# =====================================================
# CAPTURE DU SPECIFICATION TREE
# =====================================================

def capture_catia_window(output_file):

    window = activate_catia_window()

    if window is None:
        return False

    # Création automatique du dossier captures
    folder = os.path.dirname(output_file)

    if folder != "":
        os.makedirs(folder, exist_ok=True)

    # ==================================================
    # PARAMETRES A AJUSTER UNE SEULE FOIS
    # ==================================================

    # Bord gauche de la fenêtre
    LEFT_MARGIN = 5

    # Ignore le ruban CATIA (Démarrer, ENOVIA, etc.)
    TOP_MARGIN = 108

    # Largeur complète du Specification Tree
    TREE_WIDTH = 700

    # Presque toute la hauteur
    BOTTOM_MARGIN = 108

    monitor = {

        "left": window.left + LEFT_MARGIN,

        "top": window.top + TOP_MARGIN,

        "width": TREE_WIDTH,

        "height": window.height - TOP_MARGIN - BOTTOM_MARGIN

    }

    with mss.mss() as sct:

        img = sct.grab(monitor)

        mss.tools.to_png(
            img.rgb,
            img.size,
            output=output_file
        )

    print("Capture enregistrée :", output_file)

    return True