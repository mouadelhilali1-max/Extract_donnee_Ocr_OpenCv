"""Small, overlapping CATIA tree scroll steps suitable for OCR registration."""

from __future__ import annotations

import time

import pyautogui

from screen_capture import activate_catia_window


pyautogui.FAILSAFE = False
pyautogui.PAUSE = 0.01

# The old code emitted five large -150 impulses between frames, which could
# skip whole branches.  A single 110-click pulse is a practical starting
# point on the reported CATIA workstation: it normally yields 35-70 images,
# not the ~220 images caused by an excessively small step.  ``main.py`` can
# adjust this value after inspecting real frame movement.
SCROLL_CLICKS = 110
MIN_SCROLL_CLICKS = 45
MAX_SCROLL_CLICKS = 180
SCROLL_SETTLE_SECONDS = 0.24
TREE_FOCUS_X = 180
TREE_FOCUS_Y = 250


def focus_tree(*, click: bool = False, wait_seconds: float = 0.12) -> bool:
    """Activate CATIA and place the pointer over its tree without changing it."""
    window = activate_catia_window()
    if window is None:
        return False
    pyautogui.moveTo(int(window.left) + TREE_FOCUS_X, int(window.top) + TREE_FOCUS_Y, duration=0.05)
    if click:
        pyautogui.click()
    time.sleep(max(0.0, float(wait_seconds)))
    return True


def click_tree(*, wait_seconds: float = 0.12) -> bool:
    """Give CATIA tree focus through a click; kept for manual navigation."""
    return focus_tree(click=True, wait_seconds=wait_seconds)


def scroll_down(
    nb: int = 1,
    *,
    focus: bool = True,
    settle_seconds: float = SCROLL_SETTLE_SECONDS,
    clicks: int | None = None,
) -> bool:
    """Move down conservatively; consecutive screenshots retain overlap."""
    if focus and not click_tree():
        return False
    amount = SCROLL_CLICKS if clicks is None else max(MIN_SCROLL_CLICKS, min(MAX_SCROLL_CLICKS, int(clicks)))
    for _ in range(max(1, int(nb))):
        pyautogui.scroll(-amount)
        time.sleep(max(0.0, float(settle_seconds)))
    return True


def scroll_up(
    nb: int = 1,
    *,
    focus: bool = True,
    settle_seconds: float = SCROLL_SETTLE_SECONDS,
    clicks: int | None = None,
) -> bool:
    """Move upward conservatively; mainly used for manual diagnostics."""
    if focus and not click_tree():
        return False
    amount = SCROLL_CLICKS if clicks is None else max(MIN_SCROLL_CLICKS, min(MAX_SCROLL_CLICKS, int(clicks)))
    for _ in range(max(1, int(nb))):
        pyautogui.scroll(amount)
        time.sleep(max(0.0, float(settle_seconds)))
    return True


def go_to_top(*, maximum_steps: int = 90) -> bool:
    """Return to the beginning of the tree before looking for the target."""
    if not click_tree():
        return False
    for _ in range(max(1, int(maximum_steps))):
        pyautogui.scroll(SCROLL_CLICKS * 4)
        time.sleep(0.006)
    time.sleep(0.30)
    return True


def go_to_bottom(*, maximum_steps: int = 300) -> bool:
    """Diagnostic helper; the production flow stops from image stability."""
    if not click_tree():
        return False
    for _ in range(max(1, int(maximum_steps))):
        pyautogui.scroll(-SCROLL_CLICKS * 4)
        time.sleep(0.006)
    time.sleep(0.30)
    return True
