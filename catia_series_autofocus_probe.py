"""
CATIA V5 — Probe autofocus local serie -> annotation
Version 2.0 — 17/08/2026

Pourquoi V2 ?
----------------
V1 confirmait que CATIA retrouve toutes les series, mais Search renvoie parfois
deux objets : un objet de l'arbre et une geometrie associee. Lancer "Reframe
On" sur ces deux objets conserve alors la vue generale.

V2 ne lit encore aucun IT et ne cree pas d'Excel. Il valide uniquement le vrai
autofocus local qui servira au moteur final :

    1. Search de la serie connue dans l'arbre ;
    2. isolement du candidat visuel (sans l'objet arbre) ;
    3. detection de sa surbrillance orange dans le viewer ;
    4. commande CATIA "Reframe On" + clic automatique sur cette geometrie ;
    5. capture PNG apres le zoom local.

Le script ne modifie ni la piece, ni la geometrie, ni les annotations.

Entree :
    results/tree_extraction/tree_series_latest.json

Sorties :
    results/autofocus_probe_v2/captures/<ordre>_<serie>.png
    results/autofocus_probe_v2/autofocus_probe_v2_latest.json

Execution, avec CATIA ouvert sur la piece :
    .\\.venv\\Scripts\\python.exe -B .\\catia_series_autofocus_probe_v2.py
"""
from __future__ import annotations

import datetime as dt
import json
import math
import sys
import time
import traceback
from pathlib import Path
from typing import Any

try:
    import cv2
    import numpy as np
    import pythoncom
    import win32api
    import win32com.client
    import win32con
    import win32gui
    from PIL import Image
except ImportError as exc:
    raise SystemExit(
        "Dependance manquante. Dans le venv executez : "
        "pip install opencv-python numpy pywin32 Pillow"
    ) from exc


VERSION = "2.0-isolated-candidate-orange-pick"
CAT_CAPTURE_FORMAT_BMP = 4
SETTLE_SECONDS = 0.80
ORANGE_MIN_PIXELS = 12


def root_dir() -> Path:
    return Path(__file__).resolve().parent


def clean_text(value: Any) -> str:
    return " ".join(str(value or "").replace("\x00", " ").split())


def load_series(manifest_path: Path) -> list[dict[str, Any]]:
    if not manifest_path.exists():
        raise RuntimeError(
            "Manifeste absent : " + str(manifest_path) + "\n"
            "Lancez d'abord catia_functional_tolerances.py avec cette piece."
        )
    payload = json.loads(manifest_path.read_text(encoding="utf-8"))
    source = payload.get("series", [])
    seen: set[str] = set()
    rows: list[dict[str, Any]] = []
    for position, item in enumerate(source if isinstance(source, list) else [], start=1):
        code = clean_text(item.get("series_code")).upper()
        if not code or code in seen:
            continue
        seen.add(code)
        rows.append({
            "order": int(item.get("order") or position),
            "series_code": code,
            "tree_path": clean_text(item.get("tree_path")),
        })
    if not rows:
        raise RuntimeError("Le manifeste ne contient aucune serie valide.")
    return rows


def connect_catia() -> Any:
    try:
        return win32com.client.GetActiveObject("CATIA.Application")
    except Exception as exc:
        raise RuntimeError("CATIA V5 n'est pas ouvert ou n'est pas accessible.") from exc


def capture_array(viewer: Any, tmp_dir: Path, stem: str) -> np.ndarray:
    """Capture le viewer CATIA sans prendre de screenshot Windows."""
    bmp_path = tmp_dir / f".{stem}.bmp"
    png_path = tmp_dir / f".{stem}.png"
    try:
        bmp_path.unlink(missing_ok=True)
        png_path.unlink(missing_ok=True)
        viewer.CaptureToFile(CAT_CAPTURE_FORMAT_BMP, str(bmp_path.resolve()))
        deadline = time.monotonic() + 10.0
        while time.monotonic() < deadline:
            if bmp_path.exists() and bmp_path.stat().st_size > 0:
                break
            time.sleep(0.05)
        if not bmp_path.exists() or bmp_path.stat().st_size == 0:
            raise RuntimeError("CATIA n'a pas produit la capture BMP.")
        with Image.open(bmp_path) as pil_image:
            pil_image.convert("RGB").save(png_path, "PNG")
        image = cv2.imread(str(png_path), cv2.IMREAD_COLOR)
        if image is None:
            raise RuntimeError("PNG CATIA illisible par OpenCV.")
        return image
    finally:
        try:
            bmp_path.unlink(missing_ok=True)
            png_path.unlink(missing_ok=True)
        except Exception:
            pass


def save_png(image: np.ndarray, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if not cv2.imwrite(str(output_path), image):
        raise RuntimeError("Impossible d'enregistrer : " + str(output_path))


def reset_full_view(viewer: Any) -> None:
    """Retourne a une vue complete entre deux series ; aucun objet n'est modifie."""
    try:
        viewer.Reframe()
        viewer.Update()
        time.sleep(0.35)
    except Exception:
        pass


def search_candidates(selection: Any, series_code: str) -> tuple[list[dict[str, Any]], str]:
    """Retourne les objets CATIA exacts, sans accepter une serie voisine."""
    queries = (
        f"Name={series_code},all",
        f"Name={series_code}*,all",
        f"Name=*{series_code}*,all",
    )
    last_info = ""
    for query in queries:
        try:
            selection.Clear()
            selection.Search(query)
            count = int(selection.Count2)
            if count <= 0:
                continue
            candidates: list[dict[str, Any]] = []
            for index in range(1, count + 1):
                selected = selection.Item2(index)
                value = selected.Value
                try:
                    value_name = clean_text(value.Name)
                except Exception:
                    value_name = ""
                try:
                    selected_type = clean_text(selected.Type)
                except Exception:
                    selected_type = ""
                candidates.append({
                    "index": index,
                    "value": value,
                    "name": value_name,
                    "type": selected_type,
                })
            return candidates, query
        except Exception as exc:
            last_info = clean_text(exc)
    return [], last_info


def select_one(selection: Any, value: Any) -> bool:
    try:
        selection.Clear()
        selection.Add(value)
        return int(selection.Count2) == 1
    except Exception:
        return False


def orange_pick_point(image: np.ndarray) -> tuple[tuple[int, int] | None, int, tuple[int, int, int, int] | None]:
    """Trouve la geometrie actuellement surlignee par CATIA, hors arbre gauche."""
    b, g, r = cv2.split(image)
    # La surbrillance CATIA observee dans les captures est orange. Le masque ne
    # depend pas de la couleur blanche/cyan des cadres et ne sert qu'au zoom.
    mask = ((r >= 170) & (g >= 55) & (g <= 235) & (b <= 125)).astype(np.uint8) * 255
    h, w = mask.shape
    mask[:, : max(165, int(w * 0.145))] = 0  # arbre CATIA, pas le modele.
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((2, 2), np.uint8), iterations=1)
    count, labels, stats, centroids = cv2.connectedComponentsWithStats(mask, connectivity=8)
    candidates: list[tuple[int, int, int, int, int, int, float, float]] = []
    for label in range(1, count):
        x, y, bw, bh, area = [int(v) for v in stats[label]]
        if area < ORANGE_MIN_PIXELS:
            continue
        cx, cy = [float(v) for v in centroids[label]]
        candidates.append((area, label, x, y, bw, bh, cx, cy))
    if not candidates:
        return None, 0, None
    area, label, x, y, bw, bh, cx, cy = max(candidates, key=lambda item: item[0])
    ys, xs = np.where(labels == label)
    if len(xs) == 0:
        return None, area, (x, y, bw, bh)
    # Le centre d'une boite peut etre dans le vide (ligne/arc). On clique donc
    # sur un vrai pixel orange, le plus pres possible de son centroide.
    distances = (xs.astype(float) - cx) ** 2 + (ys.astype(float) - cy) ** 2
    best = int(np.argmin(distances))
    return (int(xs[best]), int(ys[best])), area, (x, y, bw, bh)


def client_screen_mapping(catia: Any, image: np.ndarray) -> tuple[int, int, float, float]:
    """Convertit les pixels CaptureToFile vers les pixels ecran du client CATIA."""
    try:
        hwnd = int(catia.HWND)
    except Exception:
        hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        raise RuntimeError("Fenetre CATIA introuvable.")
    try:
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.10)
    except Exception:
        pass
    left, top = win32gui.ClientToScreen(hwnd, (0, 0))
    client_left, client_top, client_right, client_bottom = win32gui.GetClientRect(hwnd)
    client_w = max(1, client_right - client_left)
    client_h = max(1, client_bottom - client_top)
    image_h, image_w = image.shape[:2]
    return left, top, client_w / image_w, client_h / image_h


def reframe_by_visible_pick(catia: Any, image: np.ndarray, point: tuple[int, int]) -> tuple[bool, str]:
    """Execute la commande interactive CATIA puis clique l'objet deja surligne."""
    try:
        origin_x, origin_y, scale_x, scale_y = client_screen_mapping(catia, image)
        x = origin_x + int(round(point[0] * scale_x))
        y = origin_y + int(round(point[1] * scale_y))
        catia.StartCommand("Reframe On")
        time.sleep(0.20)
        win32api.SetCursorPos((x, y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
        time.sleep(SETTLE_SECONDS)
        return True, f"clic viewer={point[0]},{point[1]} ecran={x},{y}"
    except Exception as exc:
        # Evite qu'une commande CATIA interactive reste active apres un echec.
        try:
            win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
            win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)
        except Exception:
            pass
        return False, clean_text(exc)


def local_view_score(image: np.ndarray) -> float:
    """Indicateur de densite visuelle, seulement pour choisir le bon candidat."""
    b, g, r = cv2.split(image)
    bright = (b > 140) & (g > 140) & (r > 140)
    h, w = bright.shape
    bright[:, : max(165, int(w * 0.145))] = False
    # Une vraie vue rapprochee a davantage de traits/texte utiles a l'ecran.
    return float(np.count_nonzero(bright)) / max(1.0, float(bright.size))


def main() -> int:
    root = root_dir()
    manifest_path = root / "results" / "tree_extraction" / "tree_series_latest.json"
    output_dir = root / "results" / "autofocus_probe_v2"
    captures_dir = output_dir / "captures"
    debug_dir = output_dir / "debug_selection"
    output_dir.mkdir(parents=True, exist_ok=True)
    captures_dir.mkdir(parents=True, exist_ok=True)
    debug_dir.mkdir(parents=True, exist_ok=True)

    series_rows = load_series(manifest_path)
    pythoncom.CoInitialize()
    results: list[dict[str, Any]] = []
    selection = None
    try:
        catia = connect_catia()
        document = catia.ActiveDocument
        viewer = catia.ActiveWindow.ActiveViewer
        selection = document.Selection
        try:
            viewer.Activate()
        except Exception:
            pass
        print("\nCATIA SERIES AUTOFOCUS PROBE — V2.0")
        print(f"Piece : {clean_text(getattr(document, 'Name', ''))}")
        print(f"Series du manifeste : {len(series_rows)}")
        print("Action : candidat isole -> surbrillance orange -> Reframe On -> clic -> capture.\n")

        for position, item in enumerate(series_rows, start=1):
            code = item["series_code"]
            reset_full_view(viewer)
            candidates, search_info = search_candidates(selection, code)
            candidate_report: list[dict[str, Any]] = []
            best: dict[str, Any] | None = None

            for candidate in candidates:
                if not select_one(selection, candidate["value"]):
                    candidate_report.append({
                        "index": candidate["index"],
                        "name": candidate["name"],
                        "type": candidate["type"],
                        "select_ok": False,
                    })
                    continue
                time.sleep(0.20)
                selected_image = capture_array(viewer, debug_dir, f"{position:03d}_{code}_candidate_{candidate['index']}")
                point, orange_pixels, bbox = orange_pick_point(selected_image)
                report = {
                    "index": candidate["index"],
                    "name": candidate["name"],
                    "type": candidate["type"],
                    "select_ok": True,
                    "orange_pixels": orange_pixels,
                    "pick_point": list(point) if point else None,
                    "orange_bbox": list(bbox) if bbox else None,
                }
                candidate_report.append(report)
                if point is not None and (best is None or orange_pixels > best["orange_pixels"]):
                    best = {
                        "candidate": candidate,
                        "image": selected_image,
                        "point": point,
                        "orange_pixels": orange_pixels,
                        "report": report,
                    }

            selected_ok = best is not None
            reframe_ok = False
            reframe_info = "aucun candidat visuel orange"
            captured_ok = False
            capture_path = ""
            post_score = 0.0

            if best is not None:
                select_one(selection, best["candidate"]["value"])
                reframe_ok, reframe_info = reframe_by_visible_pick(catia, best["image"], best["point"])
                try:
                    viewer.Update()
                except Exception:
                    pass
                if reframe_ok:
                    local_image = capture_array(viewer, captures_dir, f"{position:03d}_{code}_local")
                    final_path = captures_dir / f"{position:03d}_{code}.png"
                    save_png(local_image, final_path)
                    capture_path = str(final_path.resolve())
                    captured_ok = True
                    post_score = local_view_score(local_image)

            status = "OK" if selected_ok and reframe_ok and captured_ok else "ECHEC"
            results.append({
                "order": item["order"],
                "series_code": code,
                "tree_path": item["tree_path"],
                "status": status,
                "search_ok": bool(candidates),
                "search_count": len(candidates),
                "search_query_or_error": search_info,
                "selected_visual_candidate": best["candidate"]["index"] if best else None,
                "orange_pixels": best["orange_pixels"] if best else 0,
                "reframe_ok": reframe_ok,
                "reframe_info": reframe_info,
                "capture_ok": captured_ok,
                "capture_path": capture_path,
                "local_view_score": round(post_score, 6),
                "candidates": candidate_report,
            })
            print(
                f"[{position:02d}/{len(series_rows):02d}] {code} : "
                f"Search={len(candidates)} ; "
                f"candidat={'OK' if selected_ok else 'NON'} ; "
                f"Reframe={'OK' if reframe_ok else 'NON'} ; "
                f"Capture={'OK' if captured_ok else 'NON'}",
                flush=True,
            )
    except Exception as exc:
        print("\nERREUR PROBE V2 : " + clean_text(exc), file=sys.stderr)
        traceback.print_exc()
        return 2
    finally:
        if selection is not None:
            try:
                selection.Clear()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    searched = sum(1 for item in results if item["search_ok"])
    selected = sum(1 for item in results if item["selected_visual_candidate"] is not None)
    reframed = sum(1 for item in results if item["reframe_ok"])
    captured = sum(1 for item in results if item["capture_ok"])
    complete = searched == selected == reframed == captured == len(series_rows)
    output = {
        "version": VERSION,
        "generated_at": dt.datetime.now().isoformat(timespec="seconds"),
        "status": "AUTOFOCUS_LOCAL_COMPLETE" if complete else "AUTOFOCUS_LOCAL_PARTIAL",
        "manifest": str(manifest_path.resolve()),
        "capture_directory": str(captures_dir.resolve()),
        "series_total": len(series_rows),
        "search_ok": searched,
        "visual_candidate_ok": selected,
        "reframe_ok": reframed,
        "captures_ok": captured,
        "results": results,
    }
    diagnostic_path = output_dir / "autofocus_probe_v2_latest.json"
    diagnostic_path.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")

    print("\n--- RESUME FINAL ---")
    print("Statut : " + output["status"])
    print(f"Search : {searched}/{len(series_rows)}")
    print(f"Candidat visuel : {selected}/{len(series_rows)}")
    print(f"Reframe : {reframed}/{len(series_rows)}")
    print(f"Captures locales : {captured}/{len(series_rows)}")
    print("Diagnostic : " + str(diagnostic_path))
    return 0 if complete else 1


if __name__ == "__main__":
    raise SystemExit(main())
