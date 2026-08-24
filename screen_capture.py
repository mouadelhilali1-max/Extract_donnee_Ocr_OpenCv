"""Safe screen capture helpers for CATIA's visible Specification Tree."""

from __future__ import annotations

from pathlib import Path
import time

import cv2
import mss
import numpy as np
import pygetwindow as gw


# CATIA tree crop calibration.  The values remain relative to the window and
# are clamped below, so a restored or smaller CATIA window stays safe.
LEFT_MARGIN = 5
TOP_MARGIN = 108
TREE_WIDTH = 700
BOTTOM_MARGIN = 108


def get_catia_window():
    """Return the largest visible CATIA window, if there is one."""
    try:
        candidates = list(gw.getWindowsWithTitle("CATIA"))
    except Exception:
        candidates = []
    if not candidates:
        try:
            candidates = [
                window
                for window in gw.getAllWindows()
                if "catia" in str(getattr(window, "title", "")).casefold()
            ]
        except Exception:
            candidates = []

    usable = [
        window
        for window in candidates
        if int(getattr(window, "width", 0)) > 250 and int(getattr(window, "height", 0)) > 250
    ]
    return max(usable, key=lambda window: int(window.width) * int(window.height), default=None)


def activate_catia_window(wait_seconds: float = 0.18):
    """Restore CATIA and allow its tree to redraw before a capture."""
    window = get_catia_window()
    if window is None:
        return None
    try:
        if bool(getattr(window, "isMinimized", False)):
            window.restore()
        try:
            window.activate()
        except Exception:
            # Windows may deny foreground activation while CATIA is opening;
            # the capture is still possible once the window is visible.
            pass
        time.sleep(max(0.0, float(wait_seconds)))
        return window
    except Exception as error:
        print(f"[WARN] Unable to activate CATIA: {error}")
        return None


def tree_monitor(window) -> dict[str, int]:
    """Return a valid MSS rectangle for the left CATIA tree panel."""
    width = min(TREE_WIDTH, int(window.width) - LEFT_MARGIN - 4)
    height = int(window.height) - TOP_MARGIN - BOTTOM_MARGIN
    if width < 80 or height < 100:
        raise RuntimeError("CATIA window is too small. Restore or enlarge it before scanning.")
    return {
        "left": int(window.left) + LEFT_MARGIN,
        "top": int(window.top) + TOP_MARGIN,
        "width": int(width),
        "height": int(height),
    }


def grab_catia_tree(*, window=None, activate: bool = True) -> np.ndarray | None:
    """Capture the tree in memory as an OpenCV BGR image, without saving it."""
    if window is None:
        window = activate_catia_window() if activate else get_catia_window()
    if window is None:
        return None
    with mss.mss() as screen:
        raw = np.asarray(screen.grab(tree_monitor(window)))
    if raw.size == 0:
        return None
    return cv2.cvtColor(raw, cv2.COLOR_BGRA2BGR)


def save_tree_capture(image: np.ndarray, output_file: str | Path) -> bool:
    """Save an in-memory capture without changing or deleting older runs."""
    if image is None or image.size == 0:
        return False
    destination = Path(output_file)
    destination.parent.mkdir(parents=True, exist_ok=True)
    return bool(cv2.imwrite(str(destination), image))


def capture_catia_window(output_file: str | Path) -> bool:
    """Backward-compatible capture-and-save convenience wrapper."""
    image = grab_catia_tree()
    if image is None:
        return False
    saved = save_tree_capture(image, output_file)
    if saved:
        print(f"Capture saved: {output_file}")
    return saved
