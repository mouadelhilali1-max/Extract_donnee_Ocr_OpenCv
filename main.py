import os
import time
import cv2
import numpy as np

from screen_capture import (
    get_catia_window,
    capture_catia_window
)

from tree_scroller import (
    click_tree,
    scroll_down
)

# ==========================================================
# PARAMETRES
# ==========================================================

CAPTURE_FOLDER = "captures"

MAX_CAPTURES = 300

PIXEL_THRESHOLD = 500

# ==========================================================

print("=" * 60)
print("CATIA TREE SCANNER")
print("=" * 60)

os.makedirs(CAPTURE_FOLDER, exist_ok=True)

window = get_catia_window()

if window is None:

    print("CATIA n'est pas ouverte.")
    exit()

print("CATIA trouvée")
print("Titre :", window.title)
print("Position :", window.left, window.top)
print("Taille :", window.width, "x", window.height)

print("=" * 60)

print("Cliquez sur le haut du Specification Tree")
print("Puis laissez le programme travailler...")

click_tree()

time.sleep(0.2)

previous = None

for i in range(MAX_CAPTURES):

    filename = os.path.join(
        CAPTURE_FOLDER,
        f"tree_{i:03d}.png"
    )

    print(f"Capture {i:03d}")

    capture_catia_window(filename)

    current = cv2.imread(filename)

    if current is None:
        print("Erreur lecture image.")
        break

    if previous is not None:

        diff = cv2.absdiff(previous, current)

        gray = cv2.cvtColor(
            diff,
            cv2.COLOR_BGR2GRAY
        )

        _, thresh = cv2.threshold(
            gray,
            15,
            255,
            cv2.THRESH_BINARY
        )

        changed = cv2.countNonZero(thresh)

        print("Différence :", changed)

        if changed < PIXEL_THRESHOLD:

            print("\nFin de l'arbre détectée.")
            break

    previous = current.copy()

    scroll_down()

    time.sleep(0.02)

print()

print("=" * 60)
print("SCAN TERMINE")
print("Nombre de captures :", i + 1)
print("Dossier :", os.path.abspath(CAPTURE_FOLDER))
print("=" * 60)