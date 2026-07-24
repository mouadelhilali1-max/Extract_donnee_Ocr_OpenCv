import time
import pyautogui

from screen_capture import (
    get_catia_window,
    activate_catia_window
)

# Désactive le FailSafe (évite l'arrêt lorsque la souris touche un coin)
pyautogui.FAILSAFE = False

# Réduit la pause automatique entre les actions
pyautogui.PAUSE = 0.01


def click_tree():
    """
    Active CATIA puis clique dans le Specification Tree.
    """

    window = activate_catia_window()

    if window is None:
        return False

    # Coordonnées du centre de l'arbre
    x = window.left + 180
    y = window.top + 250

    pyautogui.moveTo(x, y, duration=0.05)
    pyautogui.click()

    time.sleep(0.03)

    return True


def scroll_down(nb=1):
    """
    Défile doucement vers le bas avec un recouvrement
    entre deux captures.
    """

    if not click_tree():
        return

    for _ in range(nb):

        # Plusieurs petits scrolls donnent un meilleur résultat
        # qu'un seul gros scroll.
        for _ in range(5):

            pyautogui.scroll(-150)

            time.sleep(0.01)

        time.sleep(0.02)


def scroll_up(nb=1):
    """
    Défile vers le haut.
    """

    if not click_tree():
        return

    for _ in range(nb):

        for _ in range(5):

            pyautogui.scroll(150)

            time.sleep(0.01)

        time.sleep(0.02)


def go_to_top():
    """
    Remonte complètement l'arbre.
    """

    print("Retour en haut de l'arbre...")

    click_tree()

    for _ in range(80):

        pyautogui.scroll(300)

        time.sleep(0.005)

    time.sleep(0.2)


def go_to_bottom():
    """
    Descend complètement l'arbre.
    """

    print("Descente en bas de l'arbre...")

    click_tree()

    for _ in range(300):

        pyautogui.scroll(-300)

        time.sleep(0.005)

    time.sleep(0.2)